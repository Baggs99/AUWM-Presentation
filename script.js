// I Know Who I Am - motion and nav
// Gentle, reverent interactions only. Nothing flashy.

(function () {
  const prefersReduced = window.matchMedia(
    "(prefers-reduced-motion: reduce)"
  ).matches;

  // --- Reveal sections as they scroll into view --------------------
  const revealTargets = document.querySelectorAll("section.section");

  if (!("IntersectionObserver" in window) || prefersReduced) {
    revealTargets.forEach((el) => el.classList.add("in-view"));
  } else {
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add("in-view");
            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.08, rootMargin: "0px 0px -60px 0px" }
    );
    revealTargets.forEach((el) => observer.observe(el));
  }

  // --- Active nav link ---------------------------------------------
  const navLinks = document.querySelectorAll(".site-nav a");
  const sectionsById = new Map();
  navLinks.forEach((link) => {
    const id = link.getAttribute("href").replace("#", "");
    const section = document.getElementById(id);
    if (section) sectionsById.set(id, { link, section });
  });

  if ("IntersectionObserver" in window) {
    const navObserver = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          const ref = sectionsById.get(entry.target.id);
          if (!ref) return;
          if (entry.isIntersecting) {
            navLinks.forEach((l) => l.classList.remove("active"));
            ref.link.classList.add("active");
          }
        });
      },
      { rootMargin: "-45% 0px -50% 0px", threshold: 0 }
    );
    sectionsById.forEach(({ section }) => navObserver.observe(section));
  }

  // --- Header background tightens after first scroll ---------------
  const header = document.querySelector(".site-header");
  const onScroll = () => {
    if (!header) return;
    if (window.scrollY > 80) header.classList.add("is-scrolled");
    else header.classList.remove("is-scrolled");
  };
  document.addEventListener("scroll", onScroll, { passive: true });
  onScroll();

  // --- Gentle parallax on the hero arch and stained-glass wash -----
  if (!prefersReduced) {
    const arch = document.querySelector(".hero-arch");
    const glass = document.querySelector(".hero-stainedglass");
    const halo = document.querySelector(".pt-halo");

    let ticking = false;
    const onParallax = () => {
      if (ticking) return;
      ticking = true;
      requestAnimationFrame(() => {
        const y = window.scrollY;
        if (arch && y < window.innerHeight * 1.2) {
          arch.style.transform = `translate3d(0, ${y * 0.18}px, 0)`;
        }
        if (glass && y < window.innerHeight * 1.2) {
          glass.style.transform = `translate3d(0, ${y * 0.08}px, 0)`;
        }
        if (halo) {
          const rect = halo.getBoundingClientRect();
          const visible = rect.top < window.innerHeight && rect.bottom > 0;
          if (visible) {
            const progress = 1 - rect.top / window.innerHeight;
            halo.style.transform = `translate(-50%, -50%) scale(${
              0.92 + progress * 0.1
            })`;
          }
        }
        ticking = false;
      });
    };

    document.addEventListener("scroll", onParallax, { passive: true });
    onParallax();
  }
})();
