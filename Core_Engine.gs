const CoreLock = {
  execute(callback) {
    const lock = LockService.getScriptLock();
    const success = lock.tryLock(CONFIG.LOCK_TIMEOUT);
    if (!success) {
      throw new Error("SERVER_BUSY: Server sedang memproses antrean lain.");
    }
    try {
      return callback();
    } finally {
      lock.releaseLock();
    }
  }
};

const CoreCache = {
  get(key) {
    const cached = CacheService.getScriptCache().get(key);
    return cached ? JSON.parse(cached) : null;
  },
  put(key, data, ttl = CONFIG.CACHE_TTL) {
    CacheService.getScriptCache().put(key, JSON.stringify(data), ttl);
  },
  remove(key) {
    CacheService.getScriptCache().remove(key);
  }
};
