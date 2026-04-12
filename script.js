const slides = Array.from(document.querySelectorAll(".slide"));
const progressBar = document.getElementById("progressBar");
const slideCounter = document.getElementById("slideCounter");
const slideTitle = document.getElementById("slideTitle");
const prevButton = document.getElementById("prevButton");
const nextButton = document.getElementById("nextButton");

let currentIndex = 0;

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function formatValue(value, decimals = 0, suffix = "") {
  return `${Number(value).toFixed(decimals)}${suffix}`;
}

function animateNumber(element, target, decimals = 0, suffix = "") {
  const duration = 1100;
  const start = performance.now();

  function step(now) {
    const progress = Math.min((now - start) / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3);
    const value = target * eased;
    element.textContent = formatValue(value, decimals, suffix);
    if (progress < 1) {
      requestAnimationFrame(step);
    } else {
      element.textContent = formatValue(target, decimals, suffix);
    }
  }

  requestAnimationFrame(step);
}

function resetAndAnimateBars(slide) {
  const chart = slide.querySelector(".vertical-chart");
  if (!chart) return;

  const scaleMax = Number(chart.dataset.scaleMax || "100");
  const bars = chart.querySelectorAll(".bar-fill");
  bars.forEach((bar) => {
    clearTimeout(bar._timer);
    bar.classList.remove("is-filled");
    bar.style.height = "0%";
    bar.style.opacity = "0.22";
    bar.style.transform = "translateY(18px) scaleY(0.08)";
  });

  chart.querySelectorAll(".bar-number").forEach((label) => {
    clearTimeout(label._timer);
    label.textContent = formatValue(0, Number(label.dataset.decimals || "0"), label.dataset.suffix || "");
  });

  requestAnimationFrame(() => {
    bars.forEach((bar, index) => {
      const rawTarget = Number(bar.dataset.target || "0");
      const scaledTarget = Math.min((rawTarget / scaleMax) * 100, 100);
      bar._timer = setTimeout(() => {
        bar.style.height = `${scaledTarget}%`;
        bar.style.opacity = "1";
        bar.style.transform = "translateY(0) scaleY(1)";
        bar.classList.add("is-filled");
      }, 140 * index);
    });

    chart.querySelectorAll(".bar-number").forEach((label, index) => {
      label._timer = setTimeout(() => {
        animateNumber(label, Number(label.dataset.target || "0"), Number(label.dataset.decimals || "0"), label.dataset.suffix || "");
      }, 200 + 140 * index);
    });
  });
}

function resetAndAnimateHorizontal(slide) {
  const fills = slide.querySelectorAll(".h-fill");
  fills.forEach((fill) => {
    fill.style.width = "0%";
    fill.style.transition = "width 1s cubic-bezier(0.22, 1, 0.36, 1)";
  });

  slide.querySelectorAll(".h-row strong").forEach((label) => {
    label.textContent = `0${label.dataset.suffix || ""}`;
  });

  requestAnimationFrame(() => {
    fills.forEach((fill) => {
      fill.style.width = `${Number(fill.dataset.target || "0")}%`;
    });
    slide.querySelectorAll(".h-row strong").forEach((label) => {
      animateNumber(label, Number(label.dataset.target || "0"), Number(label.dataset.decimals || "0"), label.dataset.suffix || "");
    });
  });
}

function resetAndAnimateRing(slide) {
  const circumference = 2 * Math.PI * 54;
  const rings = slide.querySelectorAll(".ring-meter");

  rings.forEach((ring) => {
    ring.style.strokeDasharray = `${circumference}`;
    ring.style.strokeDashoffset = `${circumference}`;
  });

  slide.querySelectorAll(".ring-center strong").forEach((label) => {
    label.textContent = `0${label.dataset.suffix || ""}`;
  });

  requestAnimationFrame(() => {
    rings.forEach((ring) => {
      const target = Number(ring.dataset.target || "0");
      ring.style.transition = "stroke-dashoffset 1.2s cubic-bezier(0.22, 1, 0.36, 1)";
      ring.style.strokeDashoffset = `${circumference * (1 - target / 100)}`;
    });

    slide.querySelectorAll(".ring-center strong").forEach((label) => {
      animateNumber(label, Number(label.dataset.target || "0"), Number(label.dataset.decimals || "0"), label.dataset.suffix || "");
    });
  });
}

function animateSlide(slide) {
  if (slide.querySelector(".vertical-chart")) {
    resetAndAnimateBars(slide);
  }

  if (slide.querySelector("[data-chart='clinical']")) {
    resetAndAnimateHorizontal(slide);
  }

  if (slide.querySelector("[data-chart='major-risk']")) {
    resetAndAnimateRing(slide);
  }
}

function renderSlide(index) {
  currentIndex = clamp(index, 0, slides.length - 1);

  slides.forEach((slide, slideIndex) => {
    slide.classList.toggle("is-active", slideIndex === currentIndex);
    slide.setAttribute("aria-hidden", slideIndex === currentIndex ? "false" : "true");
  });

  const activeSlide = slides[currentIndex];
  slideTitle.textContent = activeSlide.dataset.title || `Slide ${currentIndex + 1}`;
  slideCounter.textContent = `${String(currentIndex + 1).padStart(2, "0")} / ${String(slides.length).padStart(2, "0")}`;
  progressBar.style.width = `${((currentIndex + 1) / slides.length) * 100}%`;

  prevButton.disabled = currentIndex === 0;
  nextButton.disabled = currentIndex === slides.length - 1;
  prevButton.style.opacity = currentIndex === 0 ? "0.45" : "1";
  nextButton.style.opacity = currentIndex === slides.length - 1 ? "0.45" : "1";

  animateSlide(activeSlide);
}

prevButton.addEventListener("click", () => renderSlide(currentIndex - 1));
nextButton.addEventListener("click", () => renderSlide(currentIndex + 1));

document.addEventListener("keydown", (event) => {
  if (["ArrowRight", "PageDown", " "].includes(event.key)) {
    event.preventDefault();
    renderSlide(currentIndex + 1);
  }

  if (["ArrowLeft", "PageUp", "Backspace"].includes(event.key)) {
    event.preventDefault();
    renderSlide(currentIndex - 1);
  }

  if (event.key === "Home") {
    event.preventDefault();
    renderSlide(0);
  }

  if (event.key === "End") {
    event.preventDefault();
    renderSlide(slides.length - 1);
  }
});

renderSlide(0);
