# PS5 Pro stock watcher (Denmark + Sweden)

Checks authorised retailers for PlayStation 5 Pro stock every ~5 minutes on GitHub's
servers and sends a push notification to the **ntfy** phone app when one becomes
available (online stock, store stock or pre-order). One alert per change.

- **Latest results:** Actions tab → newest "PS5 Pro stock watcher" run → summary table.
- **Test the phone alert:** Actions tab → PS5 Pro stock watcher → *Run workflow* → tick "Only send a test notification".
- **Pause / stop:** Actions tab → PS5 Pro stock watcher → `...` → *Disable workflow*.
- **Settings:** `config.json` (store-only and pre-order alerts on/off).
- The ntfy topic is stored as the repository secret `NTFY_TOPIC`, so it is not visible in this public repo.
