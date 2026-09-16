(function () {
  function createWizard(root, options) {
    const opts = options || {};
    const panes = Array.from(root.querySelectorAll("[data-wizard-pane]"));
    const steps = Array.from(root.querySelectorAll("[data-wizard-step]"));
    const back = root.querySelector("[data-wizard-back]");
    const next = root.querySelector("[data-wizard-next]");
    const meta = root.querySelector("[data-wizard-meta]");
    const title = root.querySelector("[data-wizard-title]");
    const hint = root.querySelector("[data-wizard-hint]");
    const bar = root.querySelector("[data-wizard-bar]");
    let index = 0;

    function translate(key, fallback) {
      if (key && typeof window.wizardTranslate === "function") {
        return window.wizardTranslate(key);
      }
      return fallback || key || "";
    }

    function labelFor(el, attr, keyAttr, fallback) {
      if (!el) return fallback || "";
      const key = el.getAttribute(keyAttr);
      if (key) return translate(key, el.getAttribute(attr) || fallback);
      const lang = document.documentElement.lang || "ro";
      if (lang === "en") {
        return (
          el.getAttribute(attr + "-en") ||
          el.getAttribute(attr) ||
          fallback ||
          ""
        );
      }
      return el.getAttribute(attr) || fallback || "";
    }

    function show(nextIndex, flags) {
      const settings = flags || {};
      if (!panes.length) return index;
      index = Math.max(0, Math.min(panes.length - 1, nextIndex));
      panes.forEach((pane, idx) => {
        const active = idx === index;
        pane.classList.toggle("is-active", active);
        pane.hidden = !active;
        pane.setAttribute("aria-hidden", active ? "false" : "true");
      });
      steps.forEach((step, idx) => {
        step.classList.toggle("is-current", idx === index);
        step.classList.toggle("is-done", idx < index);
        step.setAttribute("aria-current", idx === index ? "step" : "false");
      });
      if (back) back.disabled = index === 0;
      const current = panes[index];
      if (title)
        title.textContent = labelFor(
          current,
          "data-title",
          "data-title-key",
          "",
        );
      if (hint)
        hint.textContent = labelFor(current, "data-hint", "data-hint-key", "");
      if (next) {
        next.disabled = index === panes.length - 1;
        if (typeof window.wizardTranslate === "function") {
          next.setAttribute(
            "data-next-label",
            translate(
              "wizardNext",
              next.getAttribute("data-next-label") || "Continue",
            ),
          );
          next.setAttribute(
            "data-last-label",
            translate(
              "wizardGoRun",
              next.getAttribute("data-last-label") || "Go to run",
            ),
          );
        }
        const nextLabel = next.getAttribute("data-next-label") || "Next";
        const lastLabel = next.getAttribute("data-last-label") || nextLabel;
        next.textContent = index >= panes.length - 2 ? lastLabel : nextLabel;
      }
      if (meta) meta.textContent = index + 1 + " / " + panes.length;
      if (bar) {
        bar.style.width = ((index + 1) / panes.length) * 100 + "%";
      }
      if (
        settings.scroll !== false &&
        typeof root.scrollIntoView === "function"
      ) {
        root.scrollIntoView({ behavior: "smooth", block: "start" });
      }
      root.dispatchEvent(
        new CustomEvent("wizard:change", { detail: { index: index } }),
      );
      return index;
    }

    if (back) back.addEventListener("click", () => show(index - 1));
    if (next) next.addEventListener("click", () => show(index + 1));
    steps.forEach((step) => {
      step.addEventListener("click", () => {
        const target = Number(step.getAttribute("data-wizard-step"));
        if (!Number.isNaN(target)) show(target);
      });
    });

    show(0, { scroll: false });
    return {
      show: show,
      get index() {
        return index;
      },
      get count() {
        return panes.length;
      },
    };
  }

  window.createWizard = createWizard;
})();
