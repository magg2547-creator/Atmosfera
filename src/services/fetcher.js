export async function fetchSheetData(config) {
  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => {
    controller.abort();
  }, config.fetchTimeout);

  try {
    const response = await fetch(config.sheetUrl, {
      signal: controller.signal,
      cache: 'no-store',
    });

    if (!response.ok) throw new Error(`HTTP ${response.status}`);

    const payload = await response.json();

    if (payload.status !== 'ok') {
      throw new Error(payload.message || 'API returned an error');
    }

    if (!Array.isArray(payload.data)) {
      throw new Error('Invalid data format');
    }

    return payload.data;
  } catch (error) {
    if (error?.name === 'AbortError') {
      throw new Error(`Request timed out after ${Math.round(config.fetchTimeout / 1000)}s`);
    }
    throw error;
  } finally {
    window.clearTimeout(timeoutId);
  }
}

export function createManualRefresh({ state, fetchSheet }) {
  return function manualRefresh() {
    if (state.fetch.isFetching) return;

    clearInterval(state.fetch.countdown);
    fetchSheet({ showLoading: true, trigger: 'manual' });
  };
}
