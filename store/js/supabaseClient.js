const Sb = (() => {
  const cfg = window.STORE_CONFIG || {};
  if (!cfg.SUPABASE_URL || cfg.SUPABASE_URL.indexOf('PASTE_YOUR') === 0) {
    return {
      ready: false,
      client: null
    };
  }
  return {
    ready: true,
    client: window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY)
  };
})();
