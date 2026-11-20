(() => {
  const bind = () => {
    const canvas = document.getElementById('speed-chart');
    if (!canvas || canvas.dataset.mapPointLinkBound || typeof L === 'undefined') return;
    canvas.dataset.mapPointLinkBound = '1';
    const target = document.getElementById('speed-chart-viewport') || canvas;
    target.addEventListener('click', (event) => {
      const state = typeof window.__rtisGetChartState === 'function' ? window.__rtisGetChartState() : null;
      if (!state?.chartHotspots?.length || !state.map) return;
      if (state.chartWasDragged) return;
      const rect = canvas.getBoundingClientRect();
      if (!rect.width || !rect.height) return;
      const scaleX = canvas.width / rect.width;
      const dpr = window.devicePixelRatio || 2;
      const x = (event.clientX - rect.left) * scaleX / dpr;
      const y = (event.clientY - rect.top) * (canvas.height / rect.height) / dpr;
      const hit = state.chartHotspots.reduce((closest, item) => {
        const distance = Math.hypot(item.x - x, item.y - y);
        const priority = item.priority || 0;
        const closestPriority = closest?.priority || 0;
        const better = !closest || priority > closestPriority || (priority === closestPriority && distance < closest.distance);
        return distance < 22 && better ? { ...item, distance } : closest;
      }, null);
      if (window.__rtisChartPointMarker) {
        state.map.removeLayer(window.__rtisChartPointMarker);
        window.__rtisChartPointMarker = null;
      }
      if (!hit?.point || hit.point.lat == null || hit.point.lon == null) return;
      const point = hit.point;
      state.map.setView([point.lat, point.lon], Math.max(state.map.getZoom(), 15), { animate: true });
      const marker = L.circleMarker([point.lat, point.lon], {
        radius: 8,
        color: '#facc15',
        weight: 3,
        fillColor: '#2563eb',
        fillOpacity: 1,
      }).addTo(state.map);
      marker.bindPopup(
        `<strong>Selected GPS point</strong><br>Speed: ${point.speed ?? '—'} km/h<br>Date/time: ${typeof mapDateTime === 'function' ? mapDateTime(point.time) : (point.time || '—')}<br>Latitude: ${Number(point.lat).toFixed(6)}<br>Longitude: ${Number(point.lon).toFixed(6)}<br>Station: ${point.station || '—'}`
      ).openPopup();
      window.__rtisChartPointMarker = marker;
    }, true);
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
  else bind();
})();
