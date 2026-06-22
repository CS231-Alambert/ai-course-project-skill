/* ═══════════════════════════════════════════
   AI Course Project Workbench — Interactions
   ═══════════════════════════════════════════ */

/* ── Theme Toggle ──────────────────────── */
(function() {
  const saved = localStorage.getItem("theme");
  const prefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  const theme = saved || (prefersDark ? "dark" : "light");
  document.documentElement.setAttribute("data-theme", theme);

  const toggle = document.getElementById("themeToggle");
  if (toggle) {
    toggle.addEventListener("click", () => {
      const current = document.documentElement.getAttribute("data-theme");
      const next = current === "dark" ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", next);
      localStorage.setItem("theme", next);
    });
  }
})();

/* ── Copy Code Buttons ─────────────────── */
(function() {
  document.querySelectorAll(".copy-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      const block = btn.closest(".code-block");
      if (!block) return;
      const code = block.querySelector("code");
      if (!code) return;
      const text = code.textContent || "";

      try {
        await navigator.clipboard.writeText(text);
        btn.textContent = "已复制 ✓";
        btn.classList.add("copied");
        setTimeout(() => {
          btn.textContent = "复制";
          btn.classList.remove("copied");
        }, 2000);
      } catch {
        /* Fallback for older browsers */
        const ta = document.createElement("textarea");
        ta.value = text;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        btn.textContent = "已复制 ✓";
        btn.classList.add("copied");
        setTimeout(() => {
          btn.textContent = "复制";
          btn.classList.remove("copied");
        }, 2000);
      }
    });
  });
})();

/* ── Scroll-Triggered Reveal ───────────── */
(function() {
  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add("visible");
          observer.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.15, rootMargin: "0px 0px -40px 0px" }
  );

  const targets = document.querySelectorAll(
    ".feature-card, .step, .eng-item, .origin-box"
  );
  targets.forEach(el => {
    el.classList.add("reveal");
    observer.observe(el);
  });
})();

/* ── Active Nav Link on Scroll ─────────── */
(function() {
  const sections = document.querySelectorAll("section[id]");
  const navLinks = document.querySelectorAll(".nav-links a");

  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          const id = entry.target.getAttribute("id");
          navLinks.forEach(link => {
            link.classList.toggle(
              "active",
              link.getAttribute("href") === `#${id}`
            );
          });
        }
      });
    },
    { rootMargin: "-40% 0px -55% 0px" }
  );

  sections.forEach(section => observer.observe(section));
})();

/* ── Smooth Scroll (polyfill for older browsers) ── */
(function() {
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener("click", function(e) {
      e.preventDefault();
      const target = document.querySelector(this.getAttribute("href"));
      if (target) {
        target.scrollIntoView({ behavior: "smooth" });
      }
    });
  });
})();

/* ── Pipeline Node Stagger Animation ───── */
(function() {
  const nodes = document.querySelectorAll(".pipe-node");
  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          nodes.forEach((node, i) => {
            node.style.transitionDelay = `${i * 100}ms`;
            node.style.opacity = "1";
            node.style.transform = "translateY(0)";
          });
          observer.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.3 }
  );

  nodes.forEach(node => {
    node.style.opacity = "0";
    node.style.transform = "translateY(12px)";
    node.style.transition = "opacity 0.5s var(--ease-out-expo), transform 0.5s var(--ease-out-expo)";
  });

  if (nodes.length) observer.observe(nodes[0]);
})();
