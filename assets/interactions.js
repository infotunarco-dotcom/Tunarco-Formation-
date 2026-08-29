// Transition animée entre les pages du site TUNARCO
document.addEventListener('DOMContentLoaded', function () {
  var overlay = document.getElementById('pageTransition');
  if (!overlay) return;

  document.querySelectorAll('a[href]').forEach(function (link) {
    var href = link.getAttribute('href');

    // Ignorer les liens externes, ancres, mailto/tel, target=_blank
    if (!href || href.startsWith('#') || href.startsWith('http') ||
        href.startsWith('mailto:') || href.startsWith('tel:') ||
        link.target === '_blank') {
      return;
    }

    link.addEventListener('click', function (e) {
      e.preventDefault();
      overlay.classList.add('leaving');
      setTimeout(function () {
        window.location.href = href;
      }, 450);
    });
  });

  // Animation de comptage pour la barre de statistiques (accueil)
  var statNums = document.querySelectorAll('.stat-num[data-target]');
  if (statNums.length) {
    var animateCount = function (el) {
      var target = parseInt(el.getAttribute('data-target'), 10) || 0;
      var suffix = el.getAttribute('data-suffix') || '';
      var duration = 1400;
      var start = null;
      var step = function (ts) {
        if (!start) start = ts;
        var progress = Math.min((ts - start) / duration, 1);
        var eased = 1 - Math.pow(1 - progress, 3);
        var value = Math.round(eased * target);
        el.textContent = value + suffix;
        if (progress < 1) requestAnimationFrame(step);
        else el.textContent = target + suffix;
      };
      requestAnimationFrame(step);
    };

    var statsBar = document.querySelector('.stats-bar');
    if (statsBar && 'IntersectionObserver' in window) {
      var observer = new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            statNums.forEach(animateCount);
            observer.disconnect();
          }
        });
      }, { threshold: 0.3 });
      observer.observe(statsBar);
    } else {
      statNums.forEach(animateCount);
    }
  }
});
