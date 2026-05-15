export function createDonutLabelPlugin(getState) {
  return {
    id: 'donutLabel',
    afterDraw(chart) {
      if (chart.canvas.id !== 'chart-donut') return;

      const state = getState();
      const { ctx, width, height } = chart;
      const hasLiveData = state.fetch.uiState === 'ready' && state.rows.length > 0;
      ctx.save();
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillStyle = 'rgba(255,255,255,.9)';
      ctx.font = '700 22px Sora, sans-serif';
      ctx.fillText(hasLiveData ? Math.round(state.current.co2) : '-', width / 2, height / 2 - 10);
      ctx.fillStyle = 'rgba(255,255,255,.45)';
      ctx.font = '400 11px Instrument Sans, sans-serif';
      ctx.fillText(hasLiveData ? 'ppm' : '', width / 2, height / 2 + 12);
      ctx.restore();
    },
  };
}
