const slides = Array.from(document.querySelectorAll(".slide"));
const progressBar = document.getElementById("progressBar");
const slideCounter = document.getElementById("slideCounter");
const slideTitle = document.getElementById("slideTitle");
const prevButton = document.getElementById("prevButton");
const nextButton = document.getElementById("nextButton");

let currentIndex = 0;

function buildOutroSlide() {
  const lastSlide = slides[slides.length - 1];
  if (!lastSlide) return;

  lastSlide.dataset.title = "마무리";
  lastSlide.classList.add("outro-slide");

  const body = lastSlide.querySelector(".section-body");
  const footer = lastSlide.querySelector(".section-footer");

  if (body) {
    body.className = "section-body outro-body";
    body.innerHTML = `
      <div class="outro-layout">
        <div class="outro-copy">
          <p class="kicker">발표 마무리</p>
          <h2 class="title outro-title">
            <span class="gradient-text">감사합니다</span>
            <span class="outro-title-sub">질문 있으시면 편하게 말씀해주세요</span>
          </h2>
          <p class="subtitle outro-subtitle">
            AI 환각을 줄이기 위해서는 더 강한 단일 응답보다, 여러 모델이 서로를 검토하고 불확실성까지 드러내는 구조가 중요하다고 보았습니다.
          </p>
          <div class="outro-chip-row">
            <span class="pill">모델 선택 가능</span>
            <span class="pill">토론 라운드 조절</span>
            <span class="pill">판정 모델</span>
            <span class="pill">불확실성 표시</span>
          </div>
        </div>
        <div class="outro-stage" aria-hidden="true">
          <div class="outro-halo outro-halo-a"></div>
          <div class="outro-halo outro-halo-b"></div>
          <div class="outro-thanks-bubble">감사합니다</div>
          <svg class="outro-character" viewBox="0 0 420 420" xmlns="http://www.w3.org/2000/svg">
            <defs>
              <linearGradient id="outroCoat" x1="0%" y1="0%" x2="100%" y2="100%">
                <stop offset="0%" stop-color="#2563eb"/>
                <stop offset="55%" stop-color="#1e3a8a"/>
                <stop offset="100%" stop-color="#0f172a"/>
              </linearGradient>
              <linearGradient id="outroHead" x1="0%" y1="0%" x2="100%" y2="100%">
                <stop offset="0%" stop-color="#ffffff"/>
                <stop offset="100%" stop-color="#dbeafe"/>
              </linearGradient>
              <radialGradient id="outroStageGlow" cx="50%" cy="40%" r="60%">
                <stop offset="0%" stop-color="rgba(59,130,246,0.26)"/>
                <stop offset="100%" stop-color="rgba(59,130,246,0)"/>
              </radialGradient>
            </defs>
            <ellipse class="outro-shadow" cx="214" cy="364" rx="92" ry="18" fill="rgba(15,23,42,0.12)"/>
            <ellipse class="outro-back-glow" cx="214" cy="214" rx="132" ry="122" fill="url(#outroStageGlow)"/>
            <g class="outro-duck-float">
              <g class="outro-bow-rig">
                <g class="outro-body-group">
                  <path d="M146 224C142 182 164 149 204 138H260C300 149 321 183 317 226L309 293C305 327 279 349 246 349H222C186 349 159 325 155 292L146 224Z" fill="url(#outroCoat)"/>
                  <path d="M204 138C178 150 162 180 161 214C180 219 196 212 209 194L204 138Z" fill="#0b1220" opacity="0.92"/>
                  <path d="M262 138C290 150 305 181 307 214C287 220 270 212 256 194L262 138Z" fill="#0b1220" opacity="0.92"/>
                  <path d="M203 189H261V312C261 331 246 345 228 345C214 345 203 333 203 318V189Z" fill="#f8fafc" opacity="0.98"/>
                  <circle cx="231" cy="222" r="7" fill="#60a5fa"/>
                  <circle cx="231" cy="251" r="7" fill="#60a5fa"/>
                  <circle cx="231" cy="280" r="7" fill="#60a5fa"/>
                </g>
                <g class="outro-left-wing">
                  <path d="M169 223C147 219 133 231 137 249C142 268 161 277 184 274L194 236C188 228 179 224 169 223Z" fill="url(#outroCoat)"/>
                </g>
                <g class="outro-right-wing">
                  <path d="M289 221C311 216 328 227 332 245C336 265 321 278 295 278L278 239C280 231 284 224 289 221Z" fill="url(#outroCoat)"/>
                </g>
                <g class="outro-head-group">
                  <circle cx="230" cy="106" r="94" fill="url(#outroHead)"/>
                  <ellipse cx="211" cy="88" rx="14" ry="8" fill="#cbd5f5" opacity="0.5"/>
                  <ellipse cx="249" cy="88" rx="14" ry="8" fill="#cbd5f5" opacity="0.5"/>
                  <circle cx="208" cy="106" r="11" fill="#111827"/>
                  <circle cx="253" cy="106" r="11" fill="#111827"/>
                  <circle cx="211" cy="102" r="3" fill="#ffffff"/>
                  <circle cx="256" cy="102" r="3" fill="#ffffff"/>
                  <ellipse cx="231" cy="141" rx="36" ry="23" fill="#f59e0b"/>
                  <path d="M198 140C212 150 247 151 264 140" stroke="#b45309" stroke-width="4" stroke-linecap="round"/>
                  <circle cx="189" cy="141" r="8" fill="#fecdd3" opacity="0.8"/>
                  <circle cx="272" cy="141" r="8" fill="#fecdd3" opacity="0.8"/>
                  <path d="M193 46C203 27 216 19 228 19C242 19 256 28 266 47" fill="none" stroke="#eff6ff" stroke-width="11" stroke-linecap="round"/>
                  <path d="M183 35C187 15 205 0 230 0C255 0 273 14 278 35L279 44H181L183 35Z" fill="#0f172a"/>
                  <rect x="173" y="26" width="113" height="27" rx="13.5" fill="#1d4ed8"/>
                  <rect x="197" y="0" width="64" height="32" rx="16" fill="#e0f2fe" stroke="#93c5fd" stroke-width="4"/>
                  <text x="229" y="21" text-anchor="middle" fill="#1d4ed8" style="font-family:Pretendard,sans-serif;font-size:20px;font-weight:900;">AI</text>
                </g>
                <g class="outro-feet-group">
                  <path d="M187 315C182 334 188 347 203 357" fill="none" stroke="#f59e0b" stroke-width="10" stroke-linecap="round"/>
                  <path d="M260 315C255 335 262 348 277 357" fill="none" stroke="#f59e0b" stroke-width="10" stroke-linecap="round"/>
                </g>
              </g>
            </g>
          </svg>
        </div>
      </div>
    `;
  }

  if (footer) {
    footer.innerHTML = "<span>발표 마무리</span><span>Thank you</span>";
  }
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function getInitialIndex() {
  const params = new URLSearchParams(window.location.search);
  const requestedSlide = Number(params.get("slide") || "1");
  if (!Number.isFinite(requestedSlide)) {
    return 0;
  }

  return clamp(Math.floor(requestedSlide) - 1, 0, slides.length - 1);
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

function resetAndAnimateMetricTable(slide) {
  const tables = slide.querySelectorAll(".metric-table");
  tables.forEach((table) => {
    const scaleMax = Number(table.dataset.scaleMax || "100");

    table.querySelectorAll("tbody tr").forEach((row) => {
      row.classList.remove("is-visible");
    });

    table.querySelectorAll(".table-fill").forEach((fill) => {
      fill.style.width = "0%";
      fill.classList.remove("is-filled");
    });

    table.querySelectorAll(".table-value").forEach((label) => {
      label.textContent = formatValue(0, Number(label.dataset.decimals || "0"), label.dataset.suffix || "");
    });

    requestAnimationFrame(() => {
      table.querySelectorAll("tbody tr").forEach((row, index) => {
        const value = row.querySelector(".table-value");
        const fill = row.querySelector(".table-fill");
        const rawTarget = Number(fill?.dataset.target || value?.dataset.target || "0");
        const scaledTarget = Math.min((rawTarget / scaleMax) * 100, 100);
        const delay = 160 * index;

        window.setTimeout(() => {
          row.classList.add("is-visible");
        }, delay);

        if (fill) {
          window.setTimeout(() => {
            fill.style.width = `${scaledTarget}%`;
            fill.classList.add("is-filled");
          }, delay + 80);
        }

        if (value) {
          window.setTimeout(() => {
            animateNumber(
              value,
              Number(value.dataset.target || "0"),
              Number(value.dataset.decimals || "0"),
              value.dataset.suffix || "",
            );
          }, delay + 120);
        }
      });
    });
  });
}

function resetAndAnimateEvidenceBars(slide) {
  const fills = slide.querySelectorAll(".evidence-bar-fill");
  if (!fills.length) return;

  fills.forEach((fill) => {
    clearTimeout(fill._timer);
    fill.style.transition = "none";
    fill.style.transform = "scaleX(0)";
    fill.style.opacity = "0.7";
  });

  requestAnimationFrame(() => {
    fills.forEach((fill, index) => {
      fill._timer = setTimeout(() => {
        fill.style.transition = "transform 1s cubic-bezier(0.22, 1, 0.36, 1), opacity 0.45s ease";
        fill.style.transform = "scaleX(1)";
        fill.style.opacity = "1";
      }, 120 * index);
    });
  });
}

function resetAndAnimateEvidenceTable(slide) {
  const tables = slide.querySelectorAll(".evidence-summary-table");
  tables.forEach((table) => {
    const scaleMax = Number(table.dataset.scaleMax || "100");

    table.querySelectorAll("tbody tr").forEach((row) => {
      row.classList.remove("is-visible");
    });

    table.querySelectorAll(".evidence-mini-fill").forEach((fill) => {
      clearTimeout(fill._timer);
      fill.style.transition = "none";
      fill.style.transform = "scaleX(0)";
      fill.style.opacity = "0.78";
    });

    requestAnimationFrame(() => {
      table.querySelectorAll("tbody tr").forEach((row, index) => {
        const value = row.querySelector(".evidence-table-value");
        const fill = row.querySelector(".evidence-mini-fill");
        const rawTarget = Number(fill?.dataset.target || value?.dataset.target || "0");
        const scaledTarget = Math.min(rawTarget / scaleMax, 1);
        const delay = 160 * index;

        window.setTimeout(() => {
          row.classList.add("is-visible");
        }, delay);

        if (fill) {
          fill._timer = window.setTimeout(() => {
            fill.style.transition = "transform 0.9s cubic-bezier(0.22, 1, 0.36, 1), opacity 0.35s ease";
            fill.style.transform = `scaleX(${scaledTarget})`;
            fill.style.opacity = "1";
          }, delay + 70);
        }
      });
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

  if (slide.querySelector(".metric-table")) {
    resetAndAnimateMetricTable(slide);
  }

  if (slide.querySelector(".evidence-bar-board")) {
    resetAndAnimateEvidenceBars(slide);
  }

  if (slide.querySelector(".evidence-summary-table")) {
    resetAndAnimateEvidenceTable(slide);
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
  document.body.classList.toggle("is-hero-active", currentIndex === 0);
  document.body.classList.toggle("is-outro-active", currentIndex === slides.length - 1);
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

buildOutroSlide();
renderSlide(getInitialIndex());
