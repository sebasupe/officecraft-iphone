// OfficeCraft para iPhone: motor de gestos táctiles de 120 fps (ProMotion) sin tirones.
//
// Diseñado específicamente para eliminar los dos grandes causantes de tirones en WASM/egui:
//   1. Pellizco (pinch-to-zoom) acelerado 100 % en GPU mediante `transform: translate3d(...) scale(...)`
//      a 120 fps sin re-renderizar WASM en cada frame; al soltar los dedos, aplica el zoom exacto
//      al motor en un único paso limpio y retira la transformación cuando el cuadro nítido está listo.
//   2. Desplazamiento vertical/horizontal agrupado a 1 evento Wheel por `requestAnimationFrame`
//      con física de inercia idéntica a UIScrollView de iOS (deceleración 0.9978^dt) y aviso `busy()`
//      con enfriamiento para que ningún sondeo periódico interrumpa el desplazamiento.
(function () {
  "use strict";

  const SLOP = 7;
  const LONG_PRESS_MS = 400;
  const DOUBLE_TAP_MS = 285;
  const DOUBLE_TAP_DIST = 36;
  const DECAY = 0.9978; // UIScrollViewDecelerationRateNormal
  const MAX_WHEEL_CHUNK = 72; // Evita inundar WASM con 20 eventos por frame
  const MIN_PINCH_SCALE = 0.35;
  const MAX_PINCH_SCALE = 4.2;

  function attach(canvas, opts) {
    let g = null;
    let fling = null;
    let hovering = false;
    let lastTap = { t: 0, x: 0, y: 0 };
    let singleTapTimer = null;
    let idleUntil = 0;
    let committingPinch = false;

    // Acumulador por frame para coalescer touchmove -> 1 WheelEvent por rAF
    let pendingDx = 0;
    let pendingDy = 0;
    let pendingX = 0;
    let pendingY = 0;
    let scrollRaf = 0;
    let pinchRaf = 0;
    let pinchPreview = null;

    const markActive = (cooldownMs) => {
      idleUntil = Math.max(idleUntil, performance.now() + (cooldownMs || 220));
    };

    const fire = {
      move(x, y, buttons) {
        canvas.dispatchEvent(
          new MouseEvent("mousemove", {
            clientX: x,
            clientY: y,
            buttons: buttons || 0,
            bubbles: true,
            cancelable: true,
          })
        );
        hovering = true;
      },
      down(x, y) {
        canvas.dispatchEvent(
          new PointerEvent("pointerdown", {
            clientX: x,
            clientY: y,
            button: 0,
            buttons: 1,
            pointerId: 1,
            pointerType: "mouse",
            isPrimary: true,
            bubbles: true,
            cancelable: true,
          })
        );
      },
      up(x, y) {
        canvas.dispatchEvent(
          new PointerEvent("pointerup", {
            clientX: x,
            clientY: y,
            button: 0,
            buttons: 0,
            pointerId: 1,
            pointerType: "mouse",
            isPrimary: true,
            bubbles: true,
            cancelable: true,
          })
        );
      },
      leave() {
        if (!hovering) return;
        hovering = false;
        canvas.dispatchEvent(new MouseEvent("mouseleave", { clientX: -1, clientY: -1 }));
      },
      wheel(x, y, dx, dy, ctrl) {
        canvas.dispatchEvent(
          new WheelEvent("wheel", {
            clientX: x,
            clientY: y,
            deltaX: dx,
            deltaY: dy,
            deltaMode: 0,
            ctrlKey: !!ctrl,
            bubbles: true,
            cancelable: true,
          })
        );
      },
    };

    function dispatchScroll(dx, dy, x, y) {
      const r = (opts.ratio && opts.ratio()) || 1;
      const px = -dx / r;
      const py = -dy / r;
      const dist = Math.hypot(px, py);
      if (dist < 0.05) return;
      const steps = Math.max(1, Math.min(4, Math.ceil(dist / MAX_WHEEL_CHUNK)));
      for (let i = 0; i < steps; i++) {
        fire.wheel(x, y, px / steps, py / steps, false);
      }
    }

    function flushScroll() {
      scrollRaf = 0;
      if (pendingDx === 0 && pendingDy === 0) return;
      const dx = pendingDx;
      const dy = pendingDy;
      const x = pendingX;
      const y = pendingY;
      pendingDx = 0;
      pendingDy = 0;
      dispatchScroll(dx, dy, x, y);
    }

    function queueScroll(dx, dy, x, y) {
      pendingDx += dx;
      pendingDy += dy;
      pendingX = x;
      pendingY = y;
      if (!scrollRaf) {
        scrollRaf = requestAnimationFrame(flushScroll);
      }
    }

    function zoomWheel(factor, x, y) {
      if (!(factor > 0) || Math.abs(Math.log(factor)) < 0.002) return;
      ensureFocus();
      fire.move(x, y);
      fire.wheel(x, y, 0, -100 * Math.log(factor), true);
    }

    // ---------- Vista previa GPU a 120 fps durante pellizco ----------
    function applyPinchTransform() {
      pinchRaf = 0;
      if (!pinchPreview) return;
      const { tx, ty, scale } = pinchPreview;
      canvas.style.transformOrigin = "0 0";
      canvas.style.transform =
        "translate3d(" + tx.toFixed(2) + "px," + ty.toFixed(2) + "px,0) scale(" + scale.toFixed(4) + ")";
    }

    function clearCanvasTransform() {
      if (pinchRaf) {
        cancelAnimationFrame(pinchRaf);
        pinchRaf = 0;
      }
      pinchPreview = null;
      canvas.style.transition = "";
      canvas.style.transform = "";
      canvas.style.transformOrigin = "";
      canvas.style.willChange = "";
    }

    function setSlidePreview(dx) {
      canvas.style.willChange = "transform";
      canvas.style.transform = "translate3d(" + (dx * 0.32).toFixed(1) + "px,0,0)";
    }

    function resetSlidePreview() {
      canvas.style.transition = "transform 180ms cubic-bezier(.22,.8,.36,1)";
      canvas.style.transform = "translate3d(0,0,0)";
      setTimeout(() => {
        if (!g) clearCanvasTransform();
      }, 190);
    }

    const onCanvas = (list) => Array.prototype.filter.call(list, (t) => t.target === canvas);
    const findTouch = (list, id) => Array.prototype.find.call(list, (t) => t.identifier === id);
    const pairOf = (a, b) => ({
      x: (a.clientX + b.clientX) / 2,
      y: (a.clientY + b.clientY) / 2,
      d: Math.max(12, Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY)),
    });

    function stopFling() {
      if (!fling) return false;
      cancelAnimationFrame(fling.raf);
      fling = null;
      return true;
    }

    function startFling(vx, vy, x, y) {
      let last = performance.now();
      fling = { raf: 0 };
      const step = (now) => {
        if (!fling) return;
        const dt = Math.min(34, Math.max(1, now - last));
        last = now;
        const k = Math.pow(DECAY, dt);
        vx *= k;
        vy *= k;
        if (Math.hypot(vx, vy) < 0.025) {
          fling = null;
          markActive(180);
          fire.leave();
          if (opts.onScrollEnd) opts.onScrollEnd();
          return;
        }
        markActive(200);
        dispatchScroll(vx * dt, vy * dt, x, y);
        fling.raf = requestAnimationFrame(step);
      };
      fling.raf = requestAnimationFrame(step);
    }

    // Velocidad suavizada con ventana ponderada (evita saltos bruscos al levantar el dedo)
    function velocity(samples, now) {
      if (samples.length < 2) return [0, 0];
      const latest = samples[samples.length - 1];
      if (now - latest[0] > 85) return [0, 0];
      let vxSum = 0;
      let vySum = 0;
      let wSum = 0;
      for (let i = samples.length - 1; i > 0; i--) {
        const b = samples[i];
        const a = samples[i - 1];
        if (latest[0] - a[0] > 110) break;
        const dt = b[0] - a[0];
        if (dt < 2 || dt > 90) continue;
        const w = Math.exp(-(latest[0] - b[0]) / 45);
        vxSum += ((b[1] - a[1]) / dt) * w;
        vySum += ((b[2] - a[2]) / dt) * w;
        wSum += w;
      }
      if (!wSum) return [0, 0];
      const cap = (v) => Math.max(-6.5, Math.min(6.5, v));
      return [cap(vxSum / wSum), cap(vySum / wSum)];
    }

    function isEguiTextAgent(el) {
      if (!el || el === document.body || el === canvas) return false;
      const isField = el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable;
      if (!isField) return false;
      // No confundir con nuestros campos nativos (barra de fórmulas, búsqueda, hojas modales)
      if (el.id && el.id.startsWith("oc-")) return false;
      if (el.closest && el.closest(".oc-sheet, .oc-formula-bar, .oc-notes-drawer, .oc-top, .oc-app-head")) return false;
      if (canvas.parentNode && canvas.parentNode.contains(el)) return true;
      // eframe inserta un <input> oculto (opacity:0, pointer-events:none) en document.body
      const st = el.style;
      if (el.id && /egui|text_agent/i.test(el.id)) return true;
      if (st && (st.pointerEvents === "none" || st.opacity === "0")) return true;
      return false;
    }

    function findTextAgent() {
      const inputs = document.querySelectorAll("input, textarea");
      for (let i = 0; i < inputs.length; i++) {
        if (isEguiTextAgent(inputs[i])) return inputs[i];
      }
      return null;
    }

    function typing() {
      return isEguiTextAgent(document.activeElement);
    }

    function ensureFocus() {
      const active = document.activeElement;
      if (active === canvas || isEguiTextAgent(active)) return;
      try {
        canvas.focus({ preventScroll: true });
      } catch (e) {}
    }

    function focusKeyboard() {
      const agent = findTextAgent();
      if (agent) {
        try {
          agent.focus({ preventScroll: true });
          return true;
        } catch (e) {}
      }
      const r = canvas.getBoundingClientRect();
      click(r.left + r.width * 0.5, r.top + r.height * 0.35);
      const agent2 = findTextAgent();
      if (agent2) {
        try {
          agent2.focus({ preventScroll: true });
          return true;
        } catch (e) {}
      }
      return false;
    }

    function click(x, y) {
      fire.move(x, y);
      fire.down(x, y);
      fire.up(x, y);
      // En Safari iOS un segundo pointerup síncrono dentro del toque abre el teclado al 1er toque
      for (let i = 0; i < 2 && !typing(); i++) fire.up(x, y);
      if (!typing()) {
        const agent = findTextAgent();
        if (agent && opts.autoKeyboard && opts.autoKeyboard()) {
          try { agent.focus({ preventScroll: true }); } catch (e) {}
        }
      }
    }

    function startPinch(a, b, mode) {
      const p = pairOf(a, b);
      const rect = canvas.getBoundingClientRect();
      const useGpu = mode !== "raton" && mode !== "raton-zoom" && !!opts.onPinchCommit;
      g = {
        kind: "pinch",
        mode,
        useGpu,
        a: a.identifier,
        b: b.identifier,
        x0: p.x,
        y0: p.y,
        d0: p.d,
        x: p.x,
        y: p.y,
        d: p.d,
        scale: 1,
        rectLeft: rect.left,
        rectTop: rect.top,
      };
      markActive(400);
      if (useGpu) {
        canvas.style.willChange = "transform";
        canvas.style.transition = "";
      } else {
        fire.move(p.x, p.y);
      }
    }

    function onStart(e) {
      if (!e.isTrusted || committingPinch) return;
      const mine = onCanvas(e.touches);
      if (!mine.length) return;
      e.preventDefault();
      e.stopPropagation();
      ensureFocus();
      const now = performance.now();
      const stopped = stopFling();
      markActive(350);

      if (mine.length >= 2) {
        clearTimeout(singleTapTimer);
        if (g && g.timer) clearTimeout(g.timer);
        if (g && (g.kind === "drag" || g.kind === "press")) fire.up(g.x, g.y);
        if (g && g.pageSwipe) clearCanvasTransform();
        startPinch(mine[0], mine[1], g ? g.mode : opts.mode());
        return;
      }
      if (g) return;
      const t = mine[0];
      g = {
        kind: "pending",
        mode: opts.mode(),
        id: t.identifier,
        x0: t.clientX,
        y0: t.clientY,
        x: t.clientX,
        y: t.clientY,
        t0: now,
        s: [[now, t.clientX, t.clientY]],
        stopped,
        lock: null,
        pageSwipe: false,
      };
      if (g.mode === "ver" && !stopped) {
        const mineG = g;
        g.timer = setTimeout(() => {
          if (g !== mineG || g.kind !== "pending") return;
          g.kind = "press";
          fire.move(g.x0, g.y0);
          fire.down(g.x0, g.y0);
          if (opts.onLongPress) opts.onLongPress(g.x0, g.y0);
        }, LONG_PRESS_MS);
      }
    }

    function onMove(e) {
      if (!e.isTrusted || !g) return;
      if (!onCanvas(e.touches).length) return;
      e.preventDefault();
      e.stopPropagation();
      markActive(350);

      if (g.kind === "pinch") {
        const a = findTouch(e.touches, g.a);
        const b = findTouch(e.touches, g.b);
        if (!a || !b) return;
        const p = pairOf(a, b);
        if (g.useGpu) {
          const rawScale = p.d / g.d0;
          const scale = Math.max(MIN_PINCH_SCALE, Math.min(MAX_PINCH_SCALE, rawScale));
          g.scale = scale;
          const ox = g.x0 - g.rectLeft;
          const oy = g.y0 - g.rectTop;
          const nx = p.x - g.rectLeft;
          const ny = p.y - g.rectTop;
          pinchPreview = {
            scale,
            tx: nx - ox * scale,
            ty: ny - oy * scale,
          };
          if (!pinchRaf) pinchRaf = requestAnimationFrame(applyPinchTransform);
        } else {
          if (g.mode !== "raton") {
            fire.move(p.x, p.y);
            zoomWheel(p.d / (g.d || p.d), p.x, p.y);
          }
          queueScroll(p.x - g.x, p.y - g.y, p.x, p.y);
        }
        g.x = p.x;
        g.y = p.y;
        g.d = p.d;
        return;
      }

      const t = findTouch(e.touches, g.id);
      if (!t) return;
      const x = t.clientX;
      const y = t.clientY;
      const now = performance.now();
      g.s.push([now, x, y]);
      if (g.s.length > 14) g.s.shift();

      if (g.kind === "pending") {
        if (Math.hypot(x - g.x0, y - g.y0) < SLOP) return;
        clearTimeout(g.timer);
        if (g.mode === "ver") {
          g.kind = "pan";
          const dx = Math.abs(x - g.x0);
          const dy = Math.abs(y - g.y0);
          const axisMode = opts.axisLock || "doc"; // "doc" | "grid" | "free"
          if (axisMode === "grid") {
            g.lock = dy > 2.8 * dx ? "y" : dx > 2.8 * dy ? "x" : null;
          } else if (axisMode === "free") {
            g.lock = null;
          } else {
            g.lock = dy > 1.75 * dx ? "y" : dx > 1.75 * dy ? "x" : null;
          }
          g.pageSwipe = !!(g.lock === "x" && opts.swipePages && opts.swipePages());
          g.x = g.x0;
          g.y = g.y0;
          if (!g.pageSwipe) fire.move(g.x0, g.y0);
        } else {
          g.kind = "drag";
          g.x = g.x0;
          g.y = g.y0;
          fire.move(g.x0, g.y0);
          fire.down(g.x0, g.y0);
        }
      }

      if (g.kind === "pan") {
        if (g.pageSwipe) {
          setSlidePreview(x - g.x0);
        } else {
          const sx = g.lock === "y" ? 0 : x - g.x;
          const sy = g.lock === "x" ? 0 : y - g.y;
          queueScroll(sx, sy, g.x0, g.y0);
        }
      } else {
        fire.move(x, y, 1);
      }
      g.x = x;
      g.y = y;
    }

    async function finishGpuPinch(was) {
      committingPinch = true;
      markActive(450);
      try {
        if (opts.onPinchCommit) {
          await opts.onPinchCommit({
            scale: was.scale,
            x0: was.x0,
            y0: was.y0,
            centerX: was.x,
            centerY: was.y,
            panX: was.x - was.x0,
            panY: was.y - was.y0,
          });
        }
        // Espera dos frames de renderizado nítido antes de retirar la vista previa GPU
        await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      } catch (err) {
        console.warn("pinch commit", err);
      } finally {
        clearCanvasTransform();
        committingPinch = false;
        markActive(200);
        fire.leave();
      }
    }

    function onEnd(e) {
      if (!e.isTrusted || !g) return;
      const ours =
        g.kind === "pinch"
          ? Array.prototype.some.call(e.changedTouches, (t) => t.identifier === g.a || t.identifier === g.b)
          : !!findTouch(e.changedTouches, g.id);
      if (!ours) return;
      e.preventDefault();
      e.stopPropagation();
      const cancelled = e.type === "touchcancel";

      if (g.kind === "pinch") {
        const was = g;
        if (was.useGpu) {
          g = null;
          if (cancelled || Math.abs(was.scale - 1) < 0.015 && Math.hypot(was.x - was.x0, was.y - was.y0) < 4) {
            clearCanvasTransform();
            fire.leave();
          } else {
            finishGpuPinch(was);
          }
          return;
        }
        const rest = onCanvas(e.touches);
        if (rest.length === 1 && !cancelled) {
          const t = rest[0];
          g = {
            kind: "pan",
            mode: was.mode,
            id: t.identifier,
            x0: t.clientX,
            y0: t.clientY,
            x: t.clientX,
            y: t.clientY,
            s: [],
            lock: null,
            noFling: true,
            fromPinch: true,
          };
          return;
        }
        g = null;
        fire.leave();
        markActive(200);
        if (opts.onPinchEnd) opts.onPinchEnd();
        return;
      }

      clearTimeout(g.timer);
      const t = findTouch(e.changedTouches, g.id);
      const x = t ? t.clientX : g.x;
      const y = t ? t.clientY : g.y;
      const was = g;
      g = null;

      if (was.kind === "pending") {
        if (!was.stopped && !cancelled) {
          const now = performance.now();
          const isDouble =
            was.mode === "ver" &&
            now - lastTap.t < DOUBLE_TAP_MS &&
            Math.hypot(was.x0 - lastTap.x, was.y0 - lastTap.y) < DOUBLE_TAP_DIST;
          if (isDouble && opts.onDoubleTap) {
            clearTimeout(singleTapTimer);
            lastTap = { t: 0, x: 0, y: 0 };
            opts.onDoubleTap(was.x0, was.y0);
            setTimeout(() => fire.leave(), 120);
            return;
          }
          click(was.x0, was.y0);
          if (opts.onTap) opts.onTap(was.x0, was.y0);
          if (was.mode === "ver") {
            lastTap = { t: now, x: was.x0, y: was.y0 };
            clearTimeout(singleTapTimer);
            const tapX = was.x0;
            const tapY = was.y0;
            if (opts.onSingleTap) {
              singleTapTimer = setTimeout(() => {
                if (!typing()) opts.onSingleTap(tapX, tapY);
              }, DOUBLE_TAP_MS);
            }
          }
        }
        fire.leave();
      } else if (was.kind === "pan") {
        flushScroll();
        if (was.pageSwipe) {
          resetSlidePreview();
          if (!cancelled) {
            const dx = x - was.x0;
            const [vx] = velocity(was.s, performance.now());
            if (Math.abs(dx) > 52 || Math.abs(vx) > 0.42) {
              if (opts.onSwipePage) opts.onSwipePage(dx < 0 ? 1 : -1);
            }
          }
          fire.leave();
          return;
        }
        const [vx, vy] = cancelled || was.noFling ? [0, 0] : velocity(was.s, performance.now());
        const fx = was.lock === "y" ? 0 : vx;
        const fy = was.lock === "x" ? 0 : vy;
        if (Math.hypot(fx, fy) > 0.22) {
          startFling(fx, fy, was.x0, was.y0);
        } else {
          markActive(180);
          fire.leave();
          if (opts.onScrollEnd) opts.onScrollEnd();
        }
        if (was.fromPinch && opts.onPinchEnd) opts.onPinchEnd();
      } else {
        fire.move(x, y, 1);
        fire.up(x, y);
        fire.leave();
        if (was.kind === "press" && opts.onSelectEnd) {
          opts.onSelectEnd(x, y, Math.hypot(x - was.x0, y - was.y0) >= SLOP);
        }
      }
    }

    function block(e) {
      if (!e.isTrusted || e.target !== canvas) return;
      if (e.type.startsWith("pointer") && e.pointerType === "mouse") return;
      e.stopPropagation();
      if (e.cancelable && !e.type.startsWith("pointer")) e.preventDefault();
    }

    const active = { passive: false, capture: true };
    window.addEventListener("touchstart", onStart, active);
    window.addEventListener("touchmove", onMove, active);
    window.addEventListener("touchend", onEnd, active);
    window.addEventListener("touchcancel", onEnd, active);
    for (const type of [
      "pointerdown",
      "pointermove",
      "pointerup",
      "pointercancel",
      "gesturestart",
      "gesturechange",
      "gestureend",
    ]) {
      window.addEventListener(type, block, active);
    }

    return {
      stop: () => {
        stopFling();
        clearTimeout(singleTapTimer);
        clearCanvasTransform();
        if (scrollRaf) {
          cancelAnimationFrame(scrollRaf);
          scrollRaf = 0;
        }
        pendingDx = 0;
        pendingDy = 0;
        fire.leave();
      },
      busy: () => !!(g || fling || committingPinch || performance.now() < idleUntil),
      typing,
      focusKeyboard,
      zoomWheel,
      scrollNow: (dx, dy) => {
        const r = canvas.getBoundingClientRect();
        dispatchScroll(dx, dy, r.left + r.width / 2, r.top + r.height / 2);
      },
    };
  }

  window.ocGestures = { attach };
})();
