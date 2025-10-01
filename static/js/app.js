// static/js/app.js

(function () {
  function wrapWithIcon(input, iconClass) {
    if (!input || input.dataset.iconWrapped === "1") return;

    // Create wrapper: <div class="input-group"> [input] <button class="btn btn-outline-secondary"><i class="bi ..."></i></button>
    const wrapper = document.createElement("div");
    wrapper.className = "input-group";

    const parent = input.parentNode;
    parent.insertBefore(wrapper, input);
    wrapper.appendChild(input);

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn-outline-secondary";
    btn.tabIndex = -1;
    btn.innerHTML = `<i class="bi ${iconClass}"></i>`;
    btn.addEventListener("click", function () {
      if (input._flatpickr) input._flatpickr.open();
      else input.focus();
    });

    wrapper.appendChild(btn);
    input.dataset.iconWrapped = "1";
  }

  function initPickers() {
    // Date-only fields (dd.mm.yyyy)
    const dateNames = [
      "birth_date",
      "arrival_date",
      "caregiver_arrival_date",
      "caregiver_departure_date"
    ];
    dateNames.forEach((name) => {
      const el = document.querySelector(`input[name="${name}"]`);
      if (el) {
        // attach flatpickr
        flatpickr(el, {
          dateFormat: "d.m.Y",
          allowInput: true,
        });
        wrapWithIcon(el, "bi-calendar-date");
      }
    });

    // Time-only fields (HH:MM, 24h)
    const timeNames = ["arrival_time"];
    timeNames.forEach((name) => {
      const el = document.querySelector(`input[name="${name}"]`);
      if (el) {
        flatpickr(el, {
          enableTime: true,
          noCalendar: true,
          dateFormat: "H:i",
          time_24hr: true,
          allowInput: true,
        });
        wrapWithIcon(el, "bi-clock");
      }
    });

    // DateTime field (dd.mm.yyyy HH:MM) — discharge
    const dt = document.querySelector(`input[name="discharge_datetime"]`);
    if (dt) {
      flatpickr(dt, {
        enableTime: true,
        dateFormat: "d.m.Y H:i",
        time_24hr: true,
        allowInput: true,
      });
      wrapWithIcon(dt, "bi-calendar2-week");
    }
  }

  document.addEventListener("DOMContentLoaded", function () {
    initPickers();

    // If caregiver fields show/hide dynamically, re-init when toggled
    document.addEventListener("change", function (e) {
      if (e.target && e.target.name === "caregiver_exists") {
        // slight delay if the UI shows fields after change
        setTimeout(initPickers, 50);
      }
    });
  });
})();
