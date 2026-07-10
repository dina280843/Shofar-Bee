/* The Shofar Bee — publish schedule.
   Each puzzle goes live on its date; the site auto-advances. Saturdays are
   skipped simply by having no entry for them (the prior day stays up).
   Add a new {day,date,file} line to extend the run. */
window.SB_SCHEDULE = [
  { day: 1,  date: '2026-06-26', file: 'day1.html'  },
  { day: 2,  date: '2026-06-28', file: 'day2.html'  },
  { day: 3,  date: '2026-06-29', file: 'day3.html'  },
  { day: 4,  date: '2026-06-30', file: 'day4.html'  },
  { day: 5,  date: '2026-07-01', file: 'day5.html'  },
  { day: 6,  date: '2026-07-02', file: 'day6.html'  },
  { day: 7,  date: '2026-07-03', file: 'day7.html'  },
  { day: 8,  date: '2026-07-05', file: 'day8.html'  },
  { day: 9,  date: '2026-07-06', file: 'day9.html'  },
  { day: 10, date: '2026-07-07', file: 'day10.html' },
  { day: 11, date: '2026-07-08', file: 'day11.html' },
  { day: 12, date: '2026-07-09', file: 'day12.html' },
  { day: 13, date: '2026-07-10', file: 'day13.html' },
  { day: 14, date: '2026-07-12', file: 'day14.html' },
  { day: 15, date: '2026-07-13', file: 'day15.html' },
  { day: 16, date: '2026-07-14', file: 'day16.html' },
  { day: 17, date: '2026-07-15', file: 'day17.html' },
  { day: 18, date: '2026-07-16', file: 'day18.html' },
  { day: 19, date: '2026-07-17', file: 'day19.html' },
  { day: 20, date: '2026-07-19', file: 'day20.html' },
  { day: 21, date: '2026-07-20', file: 'day21.html' },
  { day: 22, date: '2026-07-21', file: 'day22.html' },
  { day: 23, date: '2026-07-22', file: 'day23.html' },
  { day: 24, date: '2026-07-23', file: 'day24.html' },
  { day: 25, date: '2026-07-24', file: 'day25.html' },
  { day: 26, date: '2026-07-26', file: 'day26.html' },
  { day: 27, date: '2026-07-27', file: 'day27.html' },
  { day: 28, date: '2026-07-28', file: 'day28.html' },
  { day: 29, date: '2026-07-29', file: 'day29.html' },
  { day: 30, date: '2026-07-30', file: 'day30.html' },
  { day: 31, date: '2026-07-31', file: 'day31.html' },
  { day: 32, date: '2026-08-02', file: 'day32.html' }
];

/* The puzzle that should be live now: the latest whose date is on or before
   "today". A ?d=YYYY-MM-DD query param overrides "today" (for testing). */
window.sbActive = function () {
  var over = new URLSearchParams(location.search).get('d');
  var today = over ? new Date(over + 'T00:00:00') : new Date();
  today.setHours(0, 0, 0, 0);
  var active = window.SB_SCHEDULE[0];
  for (var i = 0; i < window.SB_SCHEDULE.length; i++) {
    if (new Date(window.SB_SCHEDULE[i].date + 'T00:00:00') <= today) active = window.SB_SCHEDULE[i];
    else break;
  }
  return active;
};

/* Fill <nav id="sbnav"> with Substack + every puzzle up to today (never reveals
   upcoming ones). currentDay gets the highlighted "primary" style. */
window.sbBuildNav = function (currentDay) {
  var el = document.getElementById('sbnav');
  if (!el) return;
  var activeDay = window.sbActive().day;
  var h = '<a href="https://theshofarbee.substack.com/" class="nav-btn" target="_blank">Substack</a>';
  for (var i = 0; i < window.SB_SCHEDULE.length; i++) {
    var s = window.SB_SCHEDULE[i];
    if (s.day > activeDay) break;
    h += '<a href="' + s.file + '" class="nav-btn' + (s.day === currentDay ? ' primary' : '') + '">Day ' + s.day + '</a>';
  }
  el.innerHTML = h;
};
