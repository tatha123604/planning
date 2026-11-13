(() => {
  const bind = () => {
    const button = document.getElementById('speed-chart-button');
    const viewport = document.querySelector('.speed-chart-y-axis');
    if (!button || !viewport || button.dataset.outlierScaleBound) return;
    button.dataset.outlierScaleBound = '1';
    button.addEventListener('click', () => {
      if (typeof points === 'undefined' || !Array.isArray(points) || !points.length) return;
      const speeds = points.map((point) => Number(point.speed) || 0);
      for (let index = 1; index < speeds.length - 1; index += 1) {
        const previous = speeds[index - 1];
        const next = speeds[index + 1];
        if (speeds[index] > Math.max(60, Math.max(previous, next) + 35) && Math.abs(previous - next) <= 25) {
          points[index].speed = Number(((previous + next) / 2).toFixed(2));
          speeds[index] = points[index].speed;
        }
      }
      const displayMax = Math.max(...speeds, 0);
      const chartMax = Math.min(130, Math.max(40, Math.ceil(displayMax / 10) * 10));
      viewport.innerHTML = Array.from({ length: chartMax / 10 + 1 }, (_, index) => `<span>${chartMax - index * 10}</span>`).join('');
      const originalMax = Math.max;
      Math.max = (...values) => values.length === 2 && values[0] === 130 ? chartMax : originalMax(...values);
    }, true);
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
  else bind();
})();
