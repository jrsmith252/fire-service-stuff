// Live analogue clock — shared by all board pages
(function () {
  var h = document.getElementById('hourHand');
  var m = document.getElementById('minuteHand');
  var s = document.getElementById('secondHand');
  if (!h || !m || !s) return;
  function setHand(el, deg) {
    el.setAttribute('transform', 'rotate(' + deg + ' 50 50)');
  }
  function tick() {
    var now = new Date();
    var sec = now.getSeconds();
    var min = now.getMinutes() + sec / 60;
    var hr  = (now.getHours() % 12) + min / 60;
    setHand(s, sec * 6);
    setHand(m, min * 6);
    setHand(h, hr * 30);
  }
  tick();
  setInterval(tick, 1000);
})();
