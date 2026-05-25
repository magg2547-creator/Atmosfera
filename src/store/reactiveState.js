function getPathValue(obj, path) {
  return path.split('.').reduce((acc, part) => acc?.[part], obj);
}

function setPathValue(obj, path, value) {
  const parts = path.split('.');
  let current = obj;
  for (let i = 0; i < parts.length - 1; i++) {
    current = current[parts[i]];
  }
  current[parts[parts.length - 1]] = value;
}

function pathMatches(subscriberPath, changedPath) {
  if (subscriberPath === changedPath) return true;
  if (changedPath.startsWith(subscriberPath + '.')) return true;
  return false;
}

export function createReactiveState(initialState) {
  let state = structuredClone(initialState);
  const subscribers = new Map();
  let batchDepth = 0;
  let pendingPaths = new Set();
  let notifyScheduled = false;

  function flushNotifications() {
    if (batchDepth > 0 || pendingPaths.size === 0) return;

    const pathsToNotify = [...pendingPaths];
    pendingPaths.clear();
    notifyScheduled = false;

    for (const [subscriberPath, callbacks] of subscribers) {
      for (const changedPath of pathsToNotify) {
        if (pathMatches(subscriberPath, changedPath)) {
          for (const cb of callbacks) {
            cb(state);
          }
          break;
        }
      }
    }
  }

  function scheduleNotify(changedPaths) {
    for (const path of changedPaths) {
      pendingPaths.add(path);
    }

    if (notifyScheduled) return;
    notifyScheduled = true;

    queueMicrotask(() => {
      notifyScheduled = false;
      flushNotifications();
    });
  }

  function batch(fn) {
    batchDepth++;
    try {
      fn();
    } finally {
      batchDepth--;
      if (batchDepth === 0) {
        flushNotifications();
      }
    }
  }

  function subscribe(path, callback) {
    if (!subscribers.has(path)) {
      subscribers.set(path, new Set());
    }
    subscribers.get(path).add(callback);

    return () => {
      const cbs = subscribers.get(path);
      if (cbs) cbs.delete(callback);
    };
  }

  function setState(updater) {
    const updates = typeof updater === 'function' ? updater(state) : updater;

    if (typeof updates === 'object' && updates !== null) {
      const changedPaths = new Set();

      function collectChanges(obj, prefix = '') {
        for (const key of Object.keys(obj)) {
          const fullPath = prefix ? `${prefix}.${key}` : key;
          const prevValue = getPathValue(state, fullPath);
          const newValue = obj[key];

          if (typeof newValue === 'object' && newValue !== null && !Array.isArray(newValue)) {
            collectChanges(newValue, fullPath);
          } else if (prevValue !== newValue) {
            setPathValue(state, fullPath, newValue);
            changedPaths.add(fullPath);
          }
        }
      }

      collectChanges(updates);

      if (changedPaths.size > 0) {
        scheduleNotify(changedPaths);
      }
    }
  }

  return { state, setState, subscribe, batch };
}
