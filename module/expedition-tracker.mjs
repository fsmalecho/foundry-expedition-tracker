import { ExpeditionTrackerHud } from './expedition-tracker-hud.mjs';
import { EXPEDITION_TRACKER_UPDATE_HOOK, registerExpeditionTrackerSettings } from './settings.mjs';

Hooks.once('init', () => {
  registerExpeditionTrackerSettings();
});

Hooks.once('ready', async () => {
  Hooks.on('canvasReady', () => {
    ExpeditionTrackerHud.sync();
  });

  Hooks.on(EXPEDITION_TRACKER_UPDATE_HOOK, () => {
    ExpeditionTrackerHud.sync();
  });

  window.addEventListener('resize', () => {
    ExpeditionTrackerHud.sync();
  });

  await ExpeditionTrackerHud.sync();
});
