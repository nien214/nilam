(function () {
  "use strict";

  let dismissedForCurrentPortrait = false;
  let backdrop;

  function init() {
    createModal();
    bindEvents();
    evaluate();
  }

  function createModal() {
    backdrop = document.createElement("div");
    backdrop.id = "orientationHint";
    backdrop.hidden = true;
    backdrop.innerHTML = `
      <div class="orientation-card" role="dialog" aria-modal="true" aria-label="Cadangan orientasi skrin">
        <h3>Cadangan Paparan</h3>
        <p>Untuk paparan terbaik, sila guna mod <strong>Landscape</strong>.</p>
        <button type="button" id="orientationHintOk">OK</button>
      </div>
    `;
    document.body.appendChild(backdrop);

    const style = document.createElement("style");
    style.textContent = `
      #orientationHint {
        position: fixed;
        inset: 0;
        background: rgba(15, 23, 42, 0.45);
        display: grid;
        place-items: center;
        padding: 1rem;
        z-index: 2000;
      }
      #orientationHint[hidden] {
        display: none !important;
      }
      #orientationHint .orientation-card {
        width: min(360px, 100%);
        background: #fff;
        border: 1px solid #d9e1e4;
        border-radius: 12px;
        padding: 0.9rem;
        box-shadow: 0 10px 30px rgba(23, 37, 42, 0.18);
      }
      #orientationHint h3 {
        margin: 0 0 0.45rem;
        font-size: 1rem;
      }
      #orientationHint p {
        margin: 0 0 0.75rem;
        font-size: 0.9rem;
      }
      #orientationHint button {
        border: 0;
        border-radius: 8px;
        background: #127475;
        color: #fff;
        font-weight: 700;
        padding: 0.45rem 0.9rem;
        cursor: pointer;
      }
      #orientationHint button:hover {
        background: #0f8b8d;
      }
    `;
    document.head.appendChild(style);

    const okBtn = backdrop.querySelector("#orientationHintOk");
    okBtn.addEventListener("click", () => {
      dismissedForCurrentPortrait = true;
      hide();
    });
  }

  function bindEvents() {
    window.addEventListener("resize", evaluate, { passive: true });
    window.addEventListener("orientationchange", evaluate, { passive: true });
  }

  function evaluate() {
    const isMobileSize = window.matchMedia("(max-width: 900px)").matches;
    const isTouchLike = window.matchMedia("(pointer: coarse)").matches;
    const isPortrait = window.matchMedia("(orientation: portrait)").matches || window.innerHeight > window.innerWidth;
    const shouldPrompt = isMobileSize && isTouchLike && isPortrait;

    if (!shouldPrompt) {
      dismissedForCurrentPortrait = false;
      hide();
      return;
    }

    if (dismissedForCurrentPortrait) {
      hide();
      return;
    }

    show();
  }

  function show() {
    if (backdrop) {
      backdrop.hidden = false;
    }
  }

  function hide() {
    if (backdrop) {
      backdrop.hidden = true;
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init, { once: true });
  } else {
    init();
  }
})();
