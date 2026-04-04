import {
  MODULE_ID,
  TRACKER_POSITION_SETTING_KEY,
  TRACKER_ENABLED_SETTING_KEY,
  TRACKER_SETTING_KEY,
  TIMED_EFFECTS_ENABLED_SETTING_KEY,
  getTimedEffectTemplates,
} from './settings.mjs';

const TRACKER_OVERLAY_ID = 'expedition-tracker-overlay';
const TRACKER_TEMPLATE_PATH = 'modules/expedition-tracker/templates/expedition-tracker.hbs';
const TRACKER_LOCALIZATION_ROOT = 'EXPEDITION_TRACKER';
const MAX_LOG_ENTRIES = 12;
const MAX_VISIBLE_TIMED_EFFECTS_PER_SLOT = 3;
const TIMED_EFFECT_TRACK_SLOT_COUNT = 12;
const CUSTOM_TIMED_EFFECT_TEMPLATE_ID = '__custom__';
const TIMED_EFFECT_TOOLTIP_ID = 'expedition-timed-effect-tooltip';
const DEFAULT_LIGHT_SOURCE_CONFIG = Object.freeze({
  isLightSource: false,
  brightLight: 0,
  dimLight: 0,
  emissionAngle: 360,
});
const EXPEDITION_JOURNAL_NAME = 'Expedition Log';
const DEFAULT_TIMEKEEPING_PRESET = 'dungeon';
const TIMEKEEPING_PRESET_CONFIGS = Object.freeze({
  dungeon: Object.freeze({
    preset: 'dungeon',
    incrementLabelSingular: 'Turn',
    incrementLabelPlural: 'Turns',
    cycleLabelSingular: 'Hour',
    cycleLabelPlural: 'Hours',
    incrementsPerCycle: 6,
  }),
  overland: Object.freeze({
    preset: 'overland',
    incrementLabelSingular: 'Watch',
    incrementLabelPlural: 'Watches',
    cycleLabelSingular: 'Day',
    cycleLabelPlural: 'Days',
    incrementsPerCycle: 6,
  }),
  custom: Object.freeze({
    preset: 'custom',
    incrementLabelSingular: 'Turn',
    incrementLabelPlural: 'Turns',
    cycleLabelSingular: 'Hour',
    cycleLabelPlural: 'Hours',
    incrementsPerCycle: 6,
  }),
});
const DEFAULT_WANDERING_CHECK = Object.freeze({
  encounterThreshold: 1,
  dieFaces: 6,
  checkEveryTurns: 1,
  autoCheck: false,
});
const DEFAULT_TIMED_EFFECT_DURATION = 6;

let trackerUpdateQueue = Promise.resolve();
let isLogCollapsed = false;
let trackerDragState = null;

function localize(key, data = {}) {
  return data && Object.keys(data).length
    ? game.i18n.format(key, data)
    : game.i18n.localize(key);
}

function localizeTracker(key, data = {}) {
  return localize(`${TRACKER_LOCALIZATION_ROOT}.${key}`, data);
}

function isExpeditionTrackerEnabled() {
  return game.user?.isGM && game.settings.get(MODULE_ID, TRACKER_ENABLED_SETTING_KEY);
}

function isTimedEffectTrackEnabled() {
  return isExpeditionTrackerEnabled() && game.settings.get(MODULE_ID, TIMED_EFFECTS_ENABLED_SETTING_KEY);
}

function getDefaultWanderingCheck() {
  return {
    ...DEFAULT_WANDERING_CHECK,
  };
}

function getDefaultTimedEffects() {
  return [];
}

function getTimekeepingPresetConfig(preset = DEFAULT_TIMEKEEPING_PRESET) {
  return TIMEKEEPING_PRESET_CONFIGS[preset] ?? TIMEKEEPING_PRESET_CONFIGS[DEFAULT_TIMEKEEPING_PRESET];
}

function getDefaultTimekeepingConfig() {
  return {
    ...getTimekeepingPresetConfig(DEFAULT_TIMEKEEPING_PRESET),
  };
}

function getDefaultTrackerState() {
  return {
    isActive: false,
    expeditionName: '',
    logToJournal: false,
    journalEntryId: '',
    journalPageId: '',
    turnCount: 0,
    lastTurnLabel: '',
    timekeeping: getDefaultTimekeepingConfig(),
    wanderingCheck: getDefaultWanderingCheck(),
    lastWanderingCheck: null,
    logEntries: [],
    timedEffects: getDefaultTimedEffects(),
  };
}

function normalizeInteger(value, fallback, { min = 1, max = Number.MAX_SAFE_INTEGER } = {}) {
  const parsed = Math.trunc(Number(value));
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(max, Math.max(min, parsed));
}

function normalizeWanderingCheckConfig(config = {}) {
  const defaults = getDefaultWanderingCheck();
  const dieFaces = normalizeInteger(config.dieFaces, defaults.dieFaces, { min: 2, max: 1000 });
  const encounterThreshold = normalizeInteger(config.encounterThreshold, defaults.encounterThreshold, { min: 1, max: dieFaces });
  const checkEveryTurns = normalizeInteger(config.checkEveryTurns, defaults.checkEveryTurns, { min: 1, max: 1000 });
  const autoCheck = Boolean(config.autoCheck);

  return {
    encounterThreshold,
    dieFaces,
    checkEveryTurns,
    autoCheck,
  };
}

function normalizeTextValue(value, fallback) {
  const normalized = String(value ?? '').trim();
  return normalized || fallback;
}

function normalizeTimekeepingConfig(config = {}) {
  const preset = String(config.preset ?? DEFAULT_TIMEKEEPING_PRESET).trim();
  const defaults = getTimekeepingPresetConfig(preset);

  return {
    preset: defaults.preset,
    incrementLabelSingular: normalizeTextValue(config.incrementLabelSingular, defaults.incrementLabelSingular),
    incrementLabelPlural: normalizeTextValue(config.incrementLabelPlural, defaults.incrementLabelPlural),
    cycleLabelSingular: normalizeTextValue(config.cycleLabelSingular, defaults.cycleLabelSingular),
    cycleLabelPlural: normalizeTextValue(config.cycleLabelPlural, defaults.cycleLabelPlural),
    incrementsPerCycle: normalizeInteger(config.incrementsPerCycle, defaults.incrementsPerCycle, { min: 1, max: 1000 }),
  };
}

function normalizeLightSourceConfig(config = {}) {
  return {
    isLightSource: Boolean(config.isLightSource),
    brightLight: normalizeInteger(config.brightLight, DEFAULT_LIGHT_SOURCE_CONFIG.brightLight, { min: 0, max: 999 }),
    dimLight: normalizeInteger(config.dimLight, DEFAULT_LIGHT_SOURCE_CONFIG.dimLight, { min: 0, max: 999 }),
    emissionAngle: normalizeInteger(config.emissionAngle, DEFAULT_LIGHT_SOURCE_CONFIG.emissionAngle, { min: 1, max: 360 }),
  };
}

function normalizeTimedEffectInstance(effect = {}) {
  const name = String(effect.name ?? '').trim();
  const iconPath = String(effect.iconPath ?? '').trim();
  const source = String(effect.source ?? '').trim();
  const sourceLabel = String(effect.sourceLabel ?? source).trim();
  const sourceTokenId = String(effect.sourceTokenId ?? '').trim();
  const sourceSceneId = String(effect.sourceSceneId ?? '').trim();
  if (!name) return null;

  const lightSource = normalizeLightSourceConfig(effect);
  const totalTurns = normalizeInteger(effect.totalTurns, DEFAULT_TIMED_EFFECT_DURATION, { min: 1, max: 999 });
  const remainingTurns = normalizeInteger(effect.remainingTurns, totalTurns, { min: 0, max: 999 });
  if (remainingTurns <= 0) return null;

  return {
    id: String(effect.id ?? '').trim() || foundry.utils.randomID(),
    templateId: String(effect.templateId ?? '').trim(),
    name,
    iconPath,
    totalTurns,
    remainingTurns,
    source,
    sourceLabel,
    sourceTokenId,
    sourceSceneId,
    createdOrder: Number(effect.createdOrder) || Date.now(),
    isPaused: Boolean(effect.isPaused),
    ...lightSource,
  };
}

function getTimedEffectFallbackLabel(name) {
  const words = String(name ?? '').trim().split(/\s+/).filter(Boolean);
  if (!words.length) return '?';
  if (words.length === 1) return words[0].slice(0, 2).toUpperCase();
  return `${words[0][0] ?? ''}${words[1][0] ?? ''}`.toUpperCase();
}

function normalizeTrackerState(state = {}) {
  const isActive = Boolean(state.isActive);
  const expeditionName = String(state.expeditionName ?? '').trim();
  const logToJournal = Boolean(state.logToJournal);
  const journalEntryId = String(state.journalEntryId ?? '').trim();
  const journalPageId = String(state.journalPageId ?? '').trim();
  const turnCount = Math.max(0, Number(state.turnCount) || 0);
  const lastTurnLabel = String(state.lastTurnLabel ?? '').trim();
  const timekeeping = normalizeTimekeepingConfig(state.timekeeping);
  const wanderingCheck = normalizeWanderingCheckConfig(state.wanderingCheck);
  const logEntries = Array.isArray(state.logEntries)
    ? state.logEntries
      .map((entry) => normalizeLogEntry(entry))
      .filter(Boolean)
      .slice(0, MAX_LOG_ENTRIES)
    : [];
  const timedEffects = Array.isArray(state.timedEffects)
    ? state.timedEffects
      .map((effect) => normalizeTimedEffectInstance(effect))
      .filter(Boolean)
    : [];

  let lastWanderingCheck = null;
  if (state.lastWanderingCheck && typeof state.lastWanderingCheck === 'object') {
    const total = Number(state.lastWanderingCheck.total);
    if (Number.isFinite(total)) {
      lastWanderingCheck = {
        total,
        turnCount: Math.max(0, Number(state.lastWanderingCheck.turnCount) || 0),
      };
    }
  }

  return {
    isActive,
    expeditionName,
    logToJournal,
    journalEntryId,
    journalPageId,
    turnCount,
    lastTurnLabel,
    timekeeping,
    wanderingCheck,
    lastWanderingCheck,
    logEntries,
    timedEffects,
  };
}

function isEffectLightSource(effect = {}) {
  return Boolean(effect.isLightSource && effect.sourceTokenId && effect.sourceSceneId);
}

function getCurrentSceneSourceTokens() {
  return (canvas?.tokens?.placeables ?? [])
    .map((token) => ({
      id: token.document.id,
      name: String(token.document?.name ?? token.actor?.name ?? '').trim(),
      controlled: Boolean(token.controlled),
    }))
    .filter((token) => token.name)
    .sort((left, right) => left.name.localeCompare(right.name));
}

function getSingleControlledTokenId() {
  const controlledTokens = getCurrentSceneSourceTokens().filter((token) => token.controlled);
  return controlledTokens.length === 1 ? controlledTokens[0].id : '';
}

function getSceneDocument(sceneId) {
  return sceneId ? game.scenes?.get(sceneId) ?? null : null;
}

function getSourceTokenDocument(effect) {
  if (!effect?.sourceTokenId || !effect?.sourceSceneId) return null;
  return getSceneDocument(effect.sourceSceneId)?.tokens.get(effect.sourceTokenId) ?? null;
}

function getSourceDisplayLabel(effect, sourceTokenName = '') {
  return String(effect?.sourceLabel ?? '').trim()
    || String(effect?.source ?? '').trim()
    || String(sourceTokenName ?? '').trim();
}

function getSceneTokenNameById(tokenId, sceneId = canvas?.scene?.id) {
  if (!tokenId || !sceneId) return '';
  const tokenDocument = getSceneDocument(sceneId)?.tokens.get(tokenId);
  return String(tokenDocument?.name ?? tokenDocument?.actor?.name ?? '').trim();
}

function getActiveLightSourceEffectsForToken(state, tokenId, sceneId) {
  return (state?.timedEffects ?? [])
    .filter((effect) => isEffectLightSource(effect) && effect.sourceTokenId === tokenId && effect.sourceSceneId === sceneId);
}

function getControllingLightSourceEffect(effects = []) {
  return [...effects].sort((left, right) => (Number(right.createdOrder) || 0) - (Number(left.createdOrder) || 0))[0] ?? null;
}

async function applyTokenLightFromEffect(tokenDocument, effect) {
  if (!tokenDocument || !effect) return;
  await tokenDocument.update({
    'light.bright': effect.brightLight,
    'light.dim': effect.dimLight,
    'light.angle': effect.emissionAngle,
  });
}

async function resetTokenLight(tokenDocument) {
  if (!tokenDocument) return;
  await tokenDocument.update({
    'light.bright': 0,
    'light.dim': 0,
    'light.angle': 360,
  });
}

async function recalculateTokenLightForToken(state, tokenId, sceneId) {
  if (!tokenId || !sceneId) return;
  const tokenDocument = getSceneDocument(sceneId)?.tokens.get(tokenId);
  if (!tokenDocument) return;

  const controllingEffect = getControllingLightSourceEffect(getActiveLightSourceEffectsForToken(state, tokenId, sceneId));
  if (controllingEffect) {
    await applyTokenLightFromEffect(tokenDocument, controllingEffect);
    return;
  }

  await resetTokenLight(tokenDocument);
}

async function recalculateTokenLightsForEffects(state, effects = []) {
  const seen = new Set();
  for (const effect of effects) {
    if (!isEffectLightSource(effect)) continue;
    const key = `${effect.sourceSceneId}:${effect.sourceTokenId}`;
    if (seen.has(key)) continue;
    seen.add(key);
    await recalculateTokenLightForToken(state, effect.sourceTokenId, effect.sourceSceneId);
  }
}

function getTrackerState() {
  const stored = game.settings.get(MODULE_ID, TRACKER_SETTING_KEY);
  return normalizeTrackerState(foundry.utils.mergeObject(getDefaultTrackerState(), stored ?? {}, { inplace: false }));
}

function normalizeLogEntry(entry = {}) {
  const type = String(entry.type ?? '').trim();
  const text = String(entry.text ?? '').trim();
  if (!type || !text) return null;

  return {
    type,
    text,
    turnCount: Math.max(0, Number(entry.turnCount) || 0),
  };
}

function appendLogEntry(entries, entry) {
  return [normalizeLogEntry(entry), ...entries]
    .filter(Boolean)
    .slice(0, MAX_LOG_ENTRIES);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function buildJournalOwnership() {
  const ownerLevel = CONST.DOCUMENT_OWNERSHIP_LEVELS?.OWNER ?? 3;
  const noneLevel = CONST.DOCUMENT_OWNERSHIP_LEVELS?.NONE ?? 0;
  const ownership = { default: noneLevel };
  for (const user of game.users ?? []) {
    if (user?.isGM) ownership[user.id] = ownerLevel;
  }
  return ownership;
}

function getExpeditionJournalName() {
  return localizeTracker('Journal.Name') || EXPEDITION_JOURNAL_NAME;
}

async function ensureExpeditionJournal() {
  const journalName = getExpeditionJournalName();
  const ownership = buildJournalOwnership();
  let journal = game.journal?.find((entry) => entry.name === journalName) ?? null;

  if (!journal) {
    journal = await JournalEntry.create({
      name: journalName,
      ownership,
      permission: ownership,
    });
    return journal;
  }

  const needsOwnershipUpdate = Object.keys(ownership).some((key) => journal.ownership?.[key] !== ownership[key])
    || Object.keys(journal.ownership ?? {}).some((key) => !(key in ownership));
  if (needsOwnershipUpdate) {
    await journal.update({
      ownership,
      permission: ownership,
    });
  }

  return journal;
}

function getNextExpeditionPageNumber(journal) {
  let highestPrefix = 0;
  for (const page of journal.pages ?? []) {
    const match = String(page?.name ?? '').match(/^\s*(\d+)\./);
    if (!match) continue;
    highestPrefix = Math.max(highestPrefix, Number(match[1]) || 0);
  }
  return highestPrefix + 1;
}

function buildExpeditionPageName(journal, expeditionName) {
  const nextNumber = getNextExpeditionPageNumber(journal);
  return `${String(nextNumber).padStart(2, '0')}. ${expeditionName}`;
}

function buildJournalPageContent(_pageName, entries = []) {
  const lines = entries.map((entry) => `<p>${escapeHtml(entry)}</p>`);
  return lines.join('\n');
}

function getLowercaseLabel(label) {
  return String(label ?? '').trim().toLowerCase();
}

function getIncrementNumberLabel(incrementCount, timekeeping) {
  return localizeTracker('Time.IncrementCountLabel', {
    label: timekeeping.incrementLabelSingular,
    count: incrementCount,
  });
}

function getCycleNumberForIncrement(incrementCount, timekeeping) {
  return incrementCount > 0
    ? Math.floor((incrementCount - 1) / timekeeping.incrementsPerCycle) + 1
    : 1;
}

function getIncrementInCycleCount(incrementCount, timekeeping) {
  return incrementCount > 0
    ? ((incrementCount - 1) % timekeeping.incrementsPerCycle) + 1
    : 0;
}

function getCycleCountLabel(cycleCount, timekeeping) {
  return localizeTracker('Time.CycleCountLabel', {
    label: timekeeping.cycleLabelSingular,
    count: cycleCount,
  });
}

function getIncrementInCycleLabel(incrementCount, timekeeping) {
  return localizeTracker('Time.IncrementInCycleLabel', {
    increment: timekeeping.incrementLabelSingular,
    cycle: timekeeping.cycleLabelSingular,
    count: getIncrementInCycleCount(incrementCount, timekeeping),
    total: timekeeping.incrementsPerCycle,
  });
}

function getCycleHeadingLabel(cycleCount, timekeeping) {
  return localizeTracker('Time.CycleHeading', {
    label: timekeeping.cycleLabelSingular,
    count: cycleCount,
  });
}

function getTimekeepingSummaryLines(timekeeping) {
  return {
    increment: localizeTracker('Time.SummaryIncrement', {
      label: timekeeping.incrementLabelSingular,
    }),
    cycle: localizeTracker('Time.SummaryCycle', {
      label: timekeeping.cycleLabelSingular,
    }),
    cadence: localizeTracker('Time.SummaryCadence', {
      count: timekeeping.incrementsPerCycle,
      increment: getLowercaseLabel(timekeeping.incrementLabelPlural),
      cycle: getLowercaseLabel(timekeeping.cycleLabelSingular),
    }),
  };
}

function buildInitialJournalPageContent(timekeeping) {
  return [
    `<h2>${escapeHtml(getCycleHeadingLabel(1, timekeeping))}</h2>`,
    `<p>${escapeHtml(localizeTracker('Journal.Started'))}</p>`,
  ].join('\n');
}

async function appendToExpeditionJournalPage({ journalEntryId, journalPageId, entry, htmlBlocks = [] }) {
  if (!journalEntryId || !journalPageId || (!entry && !htmlBlocks.length)) return;
  const journal = game.journal?.get(journalEntryId);
  const page = journal?.pages?.get(journalPageId);
  if (!page) throw new Error(localizeTracker('Journal.Unavailable'));

  const currentContent = String(page.text?.content ?? '');
  const appendedBlocks = [
    ...htmlBlocks,
    ...(entry ? [`<p>${escapeHtml(entry)}</p>`] : []),
  ].join('\n');
  const nextContent = currentContent ? `${currentContent}\n${appendedBlocks}` : appendedBlocks;
  await page.update({
    'text.content': nextContent,
  });
}

async function createExpeditionJournalPage(expeditionName, timekeeping) {
  const trimmedName = String(expeditionName ?? '').trim();
  if (!trimmedName) throw new Error(localizeTracker('Journal.NameRequired'));

  const journal = await ensureExpeditionJournal();
  const pageName = buildExpeditionPageName(journal, trimmedName);
  const createdPages = await journal.createEmbeddedDocuments('JournalEntryPage', [{
    name: pageName,
    type: 'text',
    text: {
      format: CONST.JOURNAL_ENTRY_PAGE_FORMATS?.HTML ?? 1,
      content: buildInitialJournalPageContent(timekeeping),
    },
  }]);
  const page = createdPages?.[0];
  if (!page) throw new Error(localizeTracker('Journal.Unavailable'));

  return {
    journalEntryId: journal.id,
    journalPageId: page.id,
  };
}

function getTurnActionLabel(action) {
  switch (action) {
    case 'advance':
      return localizeTracker('TurnTypes.Advance');
    case 'rest':
      return localizeTracker('TurnTypes.Rest');
    case 'search':
      return localizeTracker('TurnTypes.Search');
    case 'listen':
      return localizeTracker('TurnTypes.Listen');
    default:
      return localizeTracker('TurnTypes.Other');
  }
}

function createIncrementLogText(incrementCount, label, timekeeping) {
  return `${getIncrementNumberLabel(incrementCount, timekeeping)}: ${label}`;
}

function createIncrementLogEntry(incrementCount, label, timekeeping) {
  return {
    type: 'turn',
    text: createIncrementLogText(incrementCount, label, timekeeping),
    turnCount: incrementCount,
  };
}

function createIncrementNoteText(incrementCount, note, timekeeping) {
  return localizeTracker('Time.NoteEntry', {
    increment: getIncrementNumberLabel(incrementCount, timekeeping),
    note,
  });
}

function createIncrementNoteEntry(incrementCount, note, timekeeping) {
  return {
    type: 'note',
    text: createIncrementNoteText(incrementCount, note, timekeeping),
    turnCount: incrementCount,
  };
}

function getWanderingCheckCadenceLabel(wanderingCheck, timekeeping) {
  if (wanderingCheck.checkEveryTurns === 1) {
    return localizeTracker('Wandering.CadenceEveryIncrement', {
      label: getLowercaseLabel(timekeeping.incrementLabelSingular),
    });
  }

  return localizeTracker('Wandering.CadenceEveryIncrements', {
    count: wanderingCheck.checkEveryTurns,
    label: getLowercaseLabel(timekeeping.incrementLabelPlural),
  });
}

function getWanderingCheckRuleLabel(wanderingCheck, timekeeping) {
  return localizeTracker('Wandering.RuleSummary', {
    threshold: wanderingCheck.encounterThreshold,
    dieFaces: wanderingCheck.dieFaces,
    cadence: getWanderingCheckCadenceLabel(wanderingCheck, timekeeping),
  });
}

function getWanderingCheckAutoLabel(wanderingCheck, timekeeping) {
  return localizeTracker(
    wanderingCheck.autoCheck
      ? 'Wandering.AutoEnabled'
      : 'Wandering.AutoDisabled',
    wanderingCheck.autoCheck
      ? { cadence: getWanderingCheckCadenceLabel(wanderingCheck, timekeeping) }
      : {},
  );
}

function getWanderingCheckOutcomeLabel(wanderingCheck, total) {
  const key = total <= wanderingCheck.encounterThreshold
    ? 'Wandering.Encounter'
    : 'Wandering.NoEncounter';
  return localizeTracker(key);
}

function getWanderingCheckResultLabel(wanderingCheck, total) {
  return localizeTracker('Wandering.Result', {
    dieFaces: wanderingCheck.dieFaces,
    total,
    outcome: getWanderingCheckOutcomeLabel(wanderingCheck, total),
  });
}

function createWanderingLogText(turnCount, wanderingCheck, total, timekeeping) {
  const resultLabel = getWanderingCheckResultLabel(wanderingCheck, total);
  if (turnCount > 0) {
    return `${getIncrementNumberLabel(turnCount, timekeeping)}: ${localizeTracker('Wandering.Check')} ${resultLabel}`;
  }
  return `${localizeTracker('Wandering.Check')} ${resultLabel}`;
}

function createWanderingLogEntry(turnCount, wanderingCheck, total, timekeeping) {
  return {
    type: 'wandering',
    text: createWanderingLogText(turnCount, wanderingCheck, total, timekeeping),
    turnCount,
  };
}

function isNewCycleIncrement(turnCount, timekeeping) {
  return turnCount > 1 && ((turnCount - 1) % timekeeping.incrementsPerCycle) === 0;
}

function createCycleHeading(turnCount, timekeeping) {
  const cycleNumber = getCycleNumberForIncrement(turnCount, timekeeping);
  return `<p>&nbsp;</p>\n<h2>${escapeHtml(getCycleHeadingLabel(cycleNumber, timekeeping))}</h2>`;
}

function createTimedEffectExpiredEntry(effect) {
  const source = getSourceDisplayLabel(effect, getSceneTokenNameById(effect.sourceTokenId, effect.sourceSceneId));
  const text = source ? localizeTracker('TimedEffects.ExpiredWithSource', {
    source,
    name: effect.name,
  }) : localizeTracker('TimedEffects.Expired', {
    name: effect.name,
  });

  return {
    type: 'timed-effect-expired',
    text,
  };
}

function createTimedEffectActivatedEntry(effect) {
  const source = getSourceDisplayLabel(effect, getSceneTokenNameById(effect.sourceTokenId, effect.sourceSceneId));
  const text = source ? localizeTracker('TimedEffects.ActivatedWithSource', {
    source,
    name: effect.name,
  }) : localizeTracker('TimedEffects.Activated', {
    name: effect.name,
  });

  return {
    type: 'timed-effect-activated',
    text,
  };
}

function createTimedEffectPausedEntry(effect) {
  const source = getSourceDisplayLabel(effect, getSceneTokenNameById(effect.sourceTokenId, effect.sourceSceneId));
  const text = source ? localizeTracker('TimedEffects.PausedWithSource', {
    source,
    name: effect.name,
  }) : localizeTracker('TimedEffects.Paused', {
    name: effect.name,
  });

  return {
    type: 'timed-effect-paused',
    text,
  };
}

function createTimedEffectResumedEntry(effect) {
  const source = getSourceDisplayLabel(effect, getSceneTokenNameById(effect.sourceTokenId, effect.sourceSceneId));
  const text = source ? localizeTracker('TimedEffects.ResumedWithSource', {
    source,
    name: effect.name,
  }) : localizeTracker('TimedEffects.Resumed', {
    name: effect.name,
  });

  return {
    type: 'timed-effect-resumed',
    text,
  };
}

function getTimedEffectsSlotIndex(remainingTurns) {
  return Math.min(Math.max(remainingTurns, 1), TIMED_EFFECT_TRACK_SLOT_COUNT) - 1;
}

function buildTimedEffectTooltipHtml(effect) {
  const source = getSourceDisplayLabel(effect, getSceneTokenNameById(effect.sourceTokenId, effect.sourceSceneId));
  const rows = [
    `<div class="expedition-timed-effect-tooltip-title">${escapeHtml(effect.name)}</div>`,
    `<div class="expedition-timed-effect-tooltip-row">${escapeHtml(localizeTracker('TimedEffects.RemainingTooltip', {
      remaining: effect.remainingTurns,
      total: effect.totalTurns,
    }))}</div>`,
  ];

  if (source) {
    rows.push(`<div class="expedition-timed-effect-tooltip-row">${escapeHtml(localizeTracker('TimedEffects.SourceTooltip', {
      source,
    }))}</div>`);
  }

  if (effect.isPaused) {
    rows.push(`<div class="expedition-timed-effect-tooltip-row expedition-timed-effect-tooltip-row-status">${escapeHtml(localizeTracker('TimedEffects.PausedStatus'))}</div>`);
  }

  if (isEffectLightSource(effect)) {
    rows.push(`<div class="expedition-timed-effect-tooltip-row">${escapeHtml(localizeTracker('TimedEffects.LightTooltip', {
      bright: effect.brightLight,
      dim: effect.dimLight,
      angle: effect.emissionAngle,
    }))}</div>`);
  }

  return rows.join('');
}

function buildTimedEffectSlots(timedEffects) {
  const slots = Array.from({ length: TIMED_EFFECT_TRACK_SLOT_COUNT }, (_value, index) => ({
    index,
    effects: [],
    overflowCount: 0,
  }));

  for (const effect of timedEffects) {
    const slot = slots[getTimedEffectsSlotIndex(effect.remainingTurns)];
    const effectViewModel = {
      id: effect.id,
      name: effect.name,
      iconPath: effect.iconPath,
      fallbackLabel: getTimedEffectFallbackLabel(effect.name),
      isPaused: effect.isPaused,
      isLightSource: isEffectLightSource(effect),
      overflowLabel: effect.remainingTurns > TIMED_EFFECT_TRACK_SLOT_COUNT
        ? `${TIMED_EFFECT_TRACK_SLOT_COUNT}+`
        : '',
    };

    if (slot.effects.length < MAX_VISIBLE_TIMED_EFFECTS_PER_SLOT) {
      slot.effects.push(effectViewModel);
    } else {
      slot.overflowCount += 1;
    }
  }

  return slots;
}

function decrementTimedEffects(timedEffects = []) {
  const nextTimedEffects = [];
  const expiredTimedEffects = [];

  for (const effect of timedEffects) {
    if (effect.isPaused) {
      nextTimedEffects.push(normalizeTimedEffectInstance(effect));
      continue;
    }
    const nextEffect = normalizeTimedEffectInstance({
      ...effect,
      remainingTurns: effect.remainingTurns - 1,
    });
    if (nextEffect) nextTimedEffects.push(nextEffect);
    else expiredTimedEffects.push(effect);
  }

  return {
    nextTimedEffects,
    expiredTimedEffects,
  };
}

async function createTimedEffectExpirationWhispers(expiredTimedEffects) {
  if (!expiredTimedEffects.length) return;

  const whisperRecipients = ChatMessage.getWhisperRecipients('gm').map((user) => user.id);
  if (!whisperRecipients.length) return;

  for (const effect of expiredTimedEffects) {
    const logEntry = createTimedEffectExpiredEntry(effect);
    await ChatMessage.create({
      speaker: {
        alias: localizeTracker('Title'),
      },
      content: `<p>${escapeHtml(logEntry.text)}</p>`,
      whisper: whisperRecipients,
      sound: CONFIG.sounds.notification,
    });
  }
}

async function createTimedEffectExtinguishedMessages(expiredTimedEffects) {
  const lightSourceEffects = expiredTimedEffects.filter((effect) => isEffectLightSource(effect));
  if (!lightSourceEffects.length) return;

  for (const effect of lightSourceEffects) {
    const tokenName = getSceneTokenNameById(effect.sourceTokenId, effect.sourceSceneId)
      || getSourceDisplayLabel(effect)
      || localizeTracker('TimedEffects.UnknownSource');
    await ChatMessage.create({
      speaker: {
        alias: localizeTracker('Title'),
      },
      content: `<p>${escapeHtml(localizeTracker('TimedEffects.ExtinguishedMessage', {
        tokenName,
        name: effect.name,
      }))}</p>`,
    });
  }
}

async function removeActiveExpeditionLights(timedEffects = []) {
  await recalculateTokenLightsForEffects({ timedEffects: [] }, timedEffects.filter((effect) => isEffectLightSource(effect)));
}

function getTimedEffectById(effectId) {
  return getTrackerState().timedEffects.find((effect) => effect.id === effectId) ?? null;
}

function buildTrackerViewModel(state) {
  const turnCount = state.turnCount;
  const timekeeping = state.timekeeping;
  const cycleNumber = getCycleNumberForIncrement(turnCount, timekeeping);
  const wanderingCheckRuleLabel = getWanderingCheckRuleLabel(state.wanderingCheck, timekeeping);
  const wanderingCheckAutoLabel = getWanderingCheckAutoLabel(state.wanderingCheck, timekeeping);
  const lastWanderingCheckLabel = state.lastWanderingCheck
    ? `${getWanderingCheckResultLabel(state.wanderingCheck, state.lastWanderingCheck.total)}${state.lastWanderingCheck.turnCount > 0 ? ` (${getIncrementNumberLabel(state.lastWanderingCheck.turnCount, timekeeping)})` : ''}`
    : localizeTracker('NoWanderingCheckRecorded');
  const timedEffectsEnabled = isTimedEffectTrackEnabled();
  const timedEffects = Array.isArray(state.timedEffects) ? state.timedEffects : [];

  return {
    isActive: state.isActive,
    timedEffectsEnabled,
    showTimedEffects: state.isActive && timedEffectsEnabled,
    timedEffectSlots: timedEffectsEnabled ? buildTimedEffectSlots(timedEffects) : [],
    incrementCountLabel: getIncrementNumberLabel(turnCount, timekeeping),
    cycleCountLabel: getCycleCountLabel(cycleNumber, timekeeping),
    incrementInCycleLabel: getIncrementInCycleLabel(turnCount, timekeeping),
    lastTurnLabel: state.lastTurnLabel || localizeTracker('NoTurnsRecorded'),
    wanderingCheckRuleLabel,
    wanderingCheckAutoLabel,
    lastWanderingCheckLabel,
    logEntries: state.logEntries,
  };
}

async function buildTrackerMarkup() {
  const tracker = buildTrackerViewModel(getTrackerState());
  return foundry.applications.handlebars.renderTemplate(TRACKER_TEMPLATE_PATH, {
    tracker,
    logCollapsed: isLogCollapsed,
  });
}

function queueTrackerStateUpdate(updater) {
  trackerUpdateQueue = trackerUpdateQueue
    .then(async () => {
      const currentState = getTrackerState();
      const nextState = await updater(currentState);
      if (!nextState) return currentState;
      const normalizedState = normalizeTrackerState(nextState);
      await game.settings.set(MODULE_ID, TRACKER_SETTING_KEY, normalizedState);
      return normalizedState;
    })
    .catch((error) => {
      console.error(`${MODULE_ID} | Expedition tracker update failed`, error);
      ui.notifications.error(String(error?.message ?? error));
      return getTrackerState();
    });

  return trackerUpdateQueue;
}

function shouldAutoRunWanderingCheck(state) {
  if (!state.isActive) return false;
  if (!state.wanderingCheck?.autoCheck) return false;
  return state.turnCount > 0 && (state.turnCount % state.wanderingCheck.checkEveryTurns) === 0;
}

async function recordTurnAction(label) {
  let shouldAutoCheck = false;
  let expiredTimedEffects = [];
  await queueTrackerStateUpdate(async (state) => {
    if (!state.isActive) {
      const nextState = {
        ...getDefaultTrackerState(),
        isActive: true,
        turnCount: 1,
        lastTurnLabel: label,
        logEntries: [createIncrementLogEntry(1, label, state.timekeeping)],
      };
      shouldAutoCheck = shouldAutoRunWanderingCheck(nextState);
      if (nextState.logToJournal && nextState.journalEntryId && nextState.journalPageId) {
        await appendToExpeditionJournalPage({
          journalEntryId: nextState.journalEntryId,
          journalPageId: nextState.journalPageId,
          entry: createIncrementLogText(1, label, nextState.timekeeping),
        });
      }
      return nextState;
    }

    const nextTurnCount = state.turnCount + 1;
    const timedEffectUpdate = isTimedEffectTrackEnabled()
      ? decrementTimedEffects(state.timedEffects)
      : { nextTimedEffects: state.timedEffects, expiredTimedEffects: [] };
    expiredTimedEffects = timedEffectUpdate.expiredTimedEffects;
    let nextLogEntries = appendLogEntry(state.logEntries, createIncrementLogEntry(nextTurnCount, label, state.timekeeping));
    for (const effect of timedEffectUpdate.expiredTimedEffects) {
      nextLogEntries = appendLogEntry(nextLogEntries, createTimedEffectExpiredEntry(effect));
    }
    const nextState = {
      ...state,
      turnCount: nextTurnCount,
      lastTurnLabel: label,
      logEntries: nextLogEntries,
      timedEffects: timedEffectUpdate.nextTimedEffects,
    };
    shouldAutoCheck = shouldAutoRunWanderingCheck(nextState);
    if (nextState.logToJournal && nextState.journalEntryId && nextState.journalPageId) {
      await appendToExpeditionJournalPage({
        journalEntryId: nextState.journalEntryId,
        journalPageId: nextState.journalPageId,
        htmlBlocks: isNewCycleIncrement(nextTurnCount, nextState.timekeeping)
          ? [createCycleHeading(nextTurnCount, nextState.timekeeping)]
          : [],
        entry: createIncrementLogText(nextTurnCount, label, nextState.timekeeping),
      });
      for (const effect of timedEffectUpdate.expiredTimedEffects) {
        await appendToExpeditionJournalPage({
          journalEntryId: nextState.journalEntryId,
          journalPageId: nextState.journalPageId,
          entry: createTimedEffectExpiredEntry(effect).text,
        });
      }
    }
    return nextState;
  });

  await recalculateTokenLightsForEffects(getTrackerState(), expiredTimedEffects);
  await createTimedEffectExpirationWhispers(expiredTimedEffects);
  await createTimedEffectExtinguishedMessages(expiredTimedEffects);

  if (shouldAutoCheck) {
    await runWanderingCheck({ automatic: true });
  }
}

async function recordCurrentTurnNote(note) {
  await queueTrackerStateUpdate(async (state) => {
    if (!state.isActive || state.turnCount <= 0) {
      ui.notifications.warn(localizeTracker('Notes.Unavailable'));
      return state;
    }

    const noteEntry = createIncrementNoteEntry(state.turnCount, note, state.timekeeping);
    const nextState = {
      ...state,
      logEntries: appendLogEntry(state.logEntries, noteEntry),
    };
    if (nextState.logToJournal && nextState.journalEntryId && nextState.journalPageId) {
      await appendToExpeditionJournalPage({
        journalEntryId: nextState.journalEntryId,
        journalPageId: nextState.journalPageId,
        entry: noteEntry.text,
      });
    }
    return nextState;
  });
}

async function runWanderingCheck({ automatic = false } = {}) {
  await queueTrackerStateUpdate(async (state) => {
    const nextState = state.isActive
      ? state
      : {
        ...getDefaultTrackerState(),
        isActive: true,
      };
    const wanderingCheck = normalizeWanderingCheckConfig(nextState.wanderingCheck);
    const roll = new Roll(`1d${wanderingCheck.dieFaces}`);
    await roll.evaluate();
    await roll.toMessage({
      speaker: ChatMessage.getSpeaker(),
      flavor: localize(
        automatic
          ? `${TRACKER_LOCALIZATION_ROOT}.Wandering.AutoFlavorWithRule`
          : `${TRACKER_LOCALIZATION_ROOT}.Wandering.FlavorWithRule`,
        {
          rule: getWanderingCheckRuleLabel(wanderingCheck, nextState.timekeeping),
        },
      ),
    }, {
      rollMode: 'gmroll',
    });

    const total = Number(roll.total) || 0;
    const nextResolvedState = {
      ...nextState,
      wanderingCheck,
      lastWanderingCheck: {
        total,
        turnCount: nextState.turnCount,
      },
      logEntries: appendLogEntry(nextState.logEntries, createWanderingLogEntry(nextState.turnCount, wanderingCheck, total, nextState.timekeeping)),
    };
    if (nextResolvedState.logToJournal && nextResolvedState.journalEntryId && nextResolvedState.journalPageId) {
      const logEntry = createWanderingLogText(nextResolvedState.turnCount, wanderingCheck, total, nextResolvedState.timekeeping);
      await appendToExpeditionJournalPage({
        journalEntryId: nextResolvedState.journalEntryId,
        journalPageId: nextResolvedState.journalPageId,
        entry: logEntry,
      });
    }
    return nextResolvedState;
  });
}

function parseExpeditionConfig(form) {
  const timekeepingPreset = String(form?.elements?.timekeepingPreset?.value ?? DEFAULT_TIMEKEEPING_PRESET).trim();
  const customTimekeepingValues = {
    incrementLabelSingular: form?.elements?.timekeepingIncrementLabelSingular?.value,
    incrementLabelPlural: form?.elements?.timekeepingIncrementLabelPlural?.value,
    cycleLabelSingular: form?.elements?.timekeepingCycleLabelSingular?.value,
    cycleLabelPlural: form?.elements?.timekeepingCycleLabelPlural?.value,
    incrementsPerCycle: form?.elements?.timekeepingIncrementsPerCycle?.value,
  };
  const wanderingCheck = normalizeWanderingCheckConfig({
    encounterThreshold: form?.elements?.wanderingEncounterThreshold?.value,
    dieFaces: form?.elements?.wanderingDieFaces?.value,
    checkEveryTurns: form?.elements?.wanderingCheckEveryTurns?.value,
    autoCheck: form?.elements?.wanderingAutoCheck?.checked,
  });
  const timekeeping = normalizeTimekeepingConfig({
    preset: timekeepingPreset,
    ...customTimekeepingValues,
  });
  const customTimekeepingIncomplete = timekeepingPreset === 'custom'
    && Object.values(customTimekeepingValues).some((value) => String(value ?? '').trim() === '');

  return {
    expeditionName: String(form?.elements?.expeditionName?.value ?? '').trim(),
    logToJournal: Boolean(form?.elements?.logToJournal?.checked),
    customTimekeepingIncomplete,
    timekeeping,
    wanderingCheck,
  };
}

function getDialogApplicationForm(dialog) {
  const root = dialog?.element?.querySelector
    ? dialog.element
    : dialog?.element?.[0];
  const form = root?.querySelector?.('form.dialog-form');
  return form?.elements ? form : null;
}

function syncTimekeepingFormPresentation(form) {
  if (!form) return;

  const preset = String(form.elements?.timekeepingPreset?.value ?? DEFAULT_TIMEKEEPING_PRESET).trim();
  const timekeeping = normalizeTimekeepingConfig({ preset });
  const isCustom = timekeeping.preset === 'custom';
  const summary = form.querySelector('[data-timekeeping-summary]');
  const customFields = form.querySelector('[data-timekeeping-custom-fields]');
  const customInputs = customFields?.querySelectorAll('input') ?? [];

  if (!isCustom) {
    form.elements.timekeepingIncrementLabelSingular.value = timekeeping.incrementLabelSingular;
    form.elements.timekeepingIncrementLabelPlural.value = timekeeping.incrementLabelPlural;
    form.elements.timekeepingCycleLabelSingular.value = timekeeping.cycleLabelSingular;
    form.elements.timekeepingCycleLabelPlural.value = timekeeping.cycleLabelPlural;
    form.elements.timekeepingIncrementsPerCycle.value = String(timekeeping.incrementsPerCycle);
  }

  const summaryLines = getTimekeepingSummaryLines(isCustom
    ? normalizeTimekeepingConfig({
      preset: 'custom',
      incrementLabelSingular: form.elements.timekeepingIncrementLabelSingular.value,
      incrementLabelPlural: form.elements.timekeepingIncrementLabelPlural.value,
      cycleLabelSingular: form.elements.timekeepingCycleLabelSingular.value,
      cycleLabelPlural: form.elements.timekeepingCycleLabelPlural.value,
      incrementsPerCycle: form.elements.timekeepingIncrementsPerCycle.value,
    })
    : timekeeping);

  summary?.toggleAttribute('hidden', isCustom);
  summary?.querySelector('[data-summary-line="increment"]')?.replaceChildren(summaryLines.increment);
  summary?.querySelector('[data-summary-line="cycle"]')?.replaceChildren(summaryLines.cycle);
  summary?.querySelector('[data-summary-line="cadence"]')?.replaceChildren(summaryLines.cadence);

  customFields?.toggleAttribute('hidden', !isCustom);
  customInputs.forEach((input) => {
    input.disabled = !isCustom;
  });
}

function initializeExpeditionDialog(dialog) {
  const form = getDialogApplicationForm(dialog);
  if (!form || form.dataset.timekeepingInitialized === 'true') return;
  form.dataset.timekeepingInitialized = 'true';

  const presetSelect = form.elements?.timekeepingPreset;
  if (presetSelect) {
    presetSelect.addEventListener('change', () => syncTimekeepingFormPresentation(form));
  }

  form.querySelectorAll('[data-timekeeping-custom-fields] input').forEach((input) => {
    input.addEventListener('input', () => syncTimekeepingFormPresentation(form));
  });

  syncTimekeepingFormPresentation(form);
}

async function promptForOtherTurnLabel() {
  const content = `
    <div class="standard-form">
      <div class="form-group">
        <label>${localizeTracker('OtherTurn.PromptLabel')}</label>
        <input type="text" name="turnLabel" placeholder="${localizeTracker('OtherTurn.Placeholder')}" autofocus>
      </div>
    </div>
  `;

  let submittedValue = null;

  await new Promise((resolve) => {
    new foundry.applications.api.DialogV2({
      window: {
        title: localizeTracker('OtherTurn.Title'),
      },
      classes: ['expedition-tracker-ui'],
      content,
      buttons: [
        {
          action: 'record',
          label: localizeTracker('OtherTurn.Action'),
          default: true,
          callback: (_event, button) => {
            submittedValue = String(button.form?.elements?.turnLabel?.value ?? '').trim();
          },
        },
        {
          action: 'cancel',
          label: localizeTracker('Common.Cancel'),
        },
      ],
      submit: () => resolve(),
    }).render(true);
  });

  return submittedValue;
}

async function promptForCurrentTurnNote() {
  const content = `
    <div class="standard-form">
      <div class="form-group stacked">
        <label style="display:block; margin-bottom:0.35rem;">${localizeTracker('Notes.Label')}</label>
        <textarea name="turnNote" rows="5" placeholder="${localizeTracker('Notes.Placeholder')}" autofocus></textarea>
      </div>
    </div>
  `;

  let submittedValue = null;

  await new Promise((resolve) => {
    new foundry.applications.api.DialogV2({
      window: {
        title: localizeTracker('Notes.Title'),
      },
      classes: ['expedition-tracker-ui'],
      content,
      buttons: [
        {
          action: 'record',
          label: localizeTracker('Notes.Action'),
          default: true,
          callback: (_event, button) => {
            submittedValue = String(button.form?.elements?.turnNote?.value ?? '').trim();
          },
        },
        {
          action: 'cancel',
          label: localizeTracker('Common.Cancel'),
        },
      ],
      submit: () => resolve(),
    }).render(true);
  });

  return submittedValue;
}

function buildTimedEffectTemplateOptions(templates) {
  return [
    ...templates.map((template) => `
    <option value="${escapeHtml(template.id)}" data-default-duration="${template.defaultDurationTurns ?? ''}">
      ${escapeHtml(template.name)}
    </option>
  `),
    `<option value="${CUSTOM_TIMED_EFFECT_TEMPLATE_ID}" data-default-duration="${DEFAULT_TIMED_EFFECT_DURATION}">
      ${escapeHtml(localizeTracker('TimedEffects.AddDialog.CustomTemplateOption'))}
    </option>`,
  ].join('');
}

function buildTimedEffectSourceOptions() {
  const names = new Set();
  for (const token of canvas?.tokens?.placeables ?? []) {
    const name = String(token.document?.name ?? token.actor?.name ?? '').trim();
    if (name) names.add(name);
  }

  return [...names].sort((left, right) => left.localeCompare(right)).map((name) => `
    <option value="${escapeHtml(name)}"></option>
  `).join('');
}

function buildTimedEffectSourceTokenOptions(tokens, selectedTokenId = '') {
  const placeholderOption = `
    <option value="">${escapeHtml(localizeTracker('TimedEffects.AddDialog.SourceTokenPlaceholder'))}</option>
  `;
  const tokenOptions = tokens.map((token) => `
    <option value="${escapeHtml(token.id)}"${token.id === selectedTokenId ? ' selected' : ''}>
      ${escapeHtml(token.name)}
    </option>
  `).join('');
  return `${placeholderOption}${tokenOptions}`;
}

function syncTimedEffectLightSourceFields(form) {
  if (!form) return;

  const lightSourceEnabled = Boolean(form.elements?.timedEffectIsLightSource?.checked);
  const freeformSourceField = form.querySelector('[data-timed-effect-freeform-source-field]');
  const tokenField = form.querySelector('[data-timed-effect-source-token-field]');
  const customLabelField = form.querySelector('[data-timed-effect-source-label-field]');
  const lightFields = form.querySelector('[data-timed-effect-light-fields]');

  freeformSourceField?.toggleAttribute('hidden', lightSourceEnabled);
  tokenField?.toggleAttribute('hidden', !lightSourceEnabled);
  customLabelField?.toggleAttribute('hidden', !lightSourceEnabled);
  lightFields?.toggleAttribute('hidden', !lightSourceEnabled);

  freeformSourceField?.querySelectorAll('input').forEach((input) => {
    input.disabled = lightSourceEnabled;
  });
  tokenField?.querySelectorAll('select, input').forEach((input) => {
    input.disabled = !lightSourceEnabled;
  });
  customLabelField?.querySelectorAll('input').forEach((input) => {
    input.disabled = !lightSourceEnabled;
  });
  lightFields?.querySelectorAll('input').forEach((input) => {
    input.disabled = !lightSourceEnabled;
  });
}

function syncTimedEffectDialogDuration(form, templates) {
  if (!form) return;

  const selectedTemplateId = String(form.elements?.timedEffectTemplateId?.value ?? '').trim();
  const selectedTemplate = templates.find((template) => template.id === selectedTemplateId);
  const durationInput = form.elements?.timedEffectDurationTurns;
  if (!selectedTemplate || !durationInput) return;

  const currentDuration = String(durationInput.value ?? '').trim();
  if (!currentDuration || durationInput.dataset.isAutoDefault === 'true') {
    durationInput.value = String(selectedTemplate.defaultDurationTurns ?? DEFAULT_TIMED_EFFECT_DURATION);
    durationInput.dataset.isAutoDefault = 'true';
  }
}

function syncTimedEffectDialogLightSettings(form, templates) {
  if (!form) return;

  const selectedTemplateId = String(form.elements?.timedEffectTemplateId?.value ?? '').trim();
  const selectedTemplate = templates.find((template) => template.id === selectedTemplateId);
  const lightToggle = form.elements?.timedEffectIsLightSource;
  const brightInput = form.elements?.timedEffectBrightLight;
  const dimInput = form.elements?.timedEffectDimLight;
  const angleInput = form.elements?.timedEffectEmissionAngle;

  if (!lightToggle || !brightInput || !dimInput || !angleInput) return;

  if (!selectedTemplate) {
    if (selectedTemplateId !== CUSTOM_TIMED_EFFECT_TEMPLATE_ID) return;
    if (lightToggle.dataset.isAutoDefault === 'true') {
      lightToggle.checked = false;
    }
    if (brightInput.dataset.isAutoDefault === 'true') brightInput.value = String(DEFAULT_LIGHT_SOURCE_CONFIG.brightLight);
    if (dimInput.dataset.isAutoDefault === 'true') dimInput.value = String(DEFAULT_LIGHT_SOURCE_CONFIG.dimLight);
    if (angleInput.dataset.isAutoDefault === 'true') angleInput.value = String(DEFAULT_LIGHT_SOURCE_CONFIG.emissionAngle);
    syncTimedEffectLightSourceFields(form);
    return;
  }

  if (lightToggle.dataset.isAutoDefault === 'true') {
    lightToggle.checked = Boolean(selectedTemplate.isLightSource);
  }
  if (brightInput.dataset.isAutoDefault === 'true') {
    brightInput.value = String(selectedTemplate.brightLight ?? DEFAULT_LIGHT_SOURCE_CONFIG.brightLight);
  }
  if (dimInput.dataset.isAutoDefault === 'true') {
    dimInput.value = String(selectedTemplate.dimLight ?? DEFAULT_LIGHT_SOURCE_CONFIG.dimLight);
  }
  if (angleInput.dataset.isAutoDefault === 'true') {
    angleInput.value = String(selectedTemplate.emissionAngle ?? DEFAULT_LIGHT_SOURCE_CONFIG.emissionAngle);
  }

  syncTimedEffectLightSourceFields(form);
}

function syncTimedEffectDialogPresentation(form, templates) {
  if (!form) return;

  const selectedTemplateId = String(form.elements?.timedEffectTemplateId?.value ?? '').trim();
  const isCustom = selectedTemplateId === CUSTOM_TIMED_EFFECT_TEMPLATE_ID;
  const customFields = form.querySelector('[data-timed-effect-custom-fields]');

  customFields?.toggleAttribute('hidden', !isCustom);
  customFields?.querySelectorAll('input').forEach((input) => {
    input.disabled = !isCustom;
  });

  if (!isCustom) {
    syncTimedEffectDialogDuration(form, templates);
    syncTimedEffectDialogLightSettings(form, templates);
    return;
  }

  const durationInput = form.elements?.timedEffectDurationTurns;
  if (durationInput) {
    const currentDuration = String(durationInput.value ?? '').trim();
    if (!currentDuration || durationInput.dataset.isAutoDefault === 'true') {
      durationInput.value = String(DEFAULT_TIMED_EFFECT_DURATION);
      durationInput.dataset.isAutoDefault = 'true';
    }
  }

  syncTimedEffectDialogLightSettings(form, templates);
}

function initializeTimedEffectDialog(dialog, templates) {
  const form = getDialogApplicationForm(dialog);
  if (!form || form.dataset.timedEffectInitialized === 'true') return;
  form.dataset.timedEffectInitialized = 'true';

  const templateSelect = form.elements?.timedEffectTemplateId;
  const durationInput = form.elements?.timedEffectDurationTurns;
  const browseButton = form.querySelector('[data-action="browse-custom-timed-effect-icon"]');
  const lightToggle = form.elements?.timedEffectIsLightSource;

  templateSelect?.addEventListener('change', () => {
    if (lightToggle) lightToggle.dataset.isAutoDefault = 'true';
    form.elements?.timedEffectBrightLight?.setAttribute('data-is-auto-default', 'true');
    form.elements?.timedEffectDimLight?.setAttribute('data-is-auto-default', 'true');
    form.elements?.timedEffectEmissionAngle?.setAttribute('data-is-auto-default', 'true');
    syncTimedEffectDialogPresentation(form, templates);
  });
  durationInput?.addEventListener('input', () => {
    durationInput.dataset.isAutoDefault = 'false';
  });
  lightToggle?.addEventListener('change', () => {
    lightToggle.dataset.isAutoDefault = 'false';
    syncTimedEffectLightSourceFields(form);
  });
  form.elements?.timedEffectBrightLight?.addEventListener('input', () => {
    form.elements.timedEffectBrightLight.dataset.isAutoDefault = 'false';
  });
  form.elements?.timedEffectDimLight?.addEventListener('input', () => {
    form.elements.timedEffectDimLight.dataset.isAutoDefault = 'false';
  });
  form.elements?.timedEffectEmissionAngle?.addEventListener('input', () => {
    form.elements.timedEffectEmissionAngle.dataset.isAutoDefault = 'false';
  });

  browseButton?.addEventListener('click', (event) => {
    event.preventDefault();
    const input = form.elements?.timedEffectCustomIconPath;
    if (!input) return;

    new FilePicker({
      type: 'imagevideo',
      current: String(input.value ?? '').trim(),
      callback: (path) => {
        input.value = path;
      },
    }).browse();
  });

  syncTimedEffectDialogPresentation(form, templates);
}

function parseTimedEffectConfig(form, templates) {
  const templateId = String(form?.elements?.timedEffectTemplateId?.value ?? '').trim();
  const sourceTokenId = String(form?.elements?.timedEffectSourceTokenId?.value ?? '').trim();
  const sourceSceneId = String(canvas?.scene?.id ?? '').trim();
  const sourceTokenName = getSceneTokenNameById(sourceTokenId, sourceSceneId);
  const sourceLabel = String(form?.elements?.timedEffectSourceLabel?.value ?? '').trim();
  const freeformSource = String(form?.elements?.timedEffectSource?.value ?? '').trim();
  const lightSourceConfig = normalizeLightSourceConfig({
    isLightSource: form?.elements?.timedEffectIsLightSource?.checked,
    brightLight: form?.elements?.timedEffectBrightLight?.value,
    dimLight: form?.elements?.timedEffectDimLight?.value,
    emissionAngle: form?.elements?.timedEffectEmissionAngle?.value,
  });

  if (lightSourceConfig.isLightSource && (!sourceTokenId || !sourceSceneId || !sourceTokenName)) {
    ui.notifications.warn(localizeTracker('TimedEffects.AddDialog.SourceTokenRequired'));
    return null;
  }

  if (templateId === CUSTOM_TIMED_EFFECT_TEMPLATE_ID) {
    const name = String(form?.elements?.timedEffectCustomName?.value ?? '').trim();
    const iconPath = String(form?.elements?.timedEffectCustomIconPath?.value ?? '').trim();
    if (!name) {
      ui.notifications.warn(localizeTracker('TimedEffects.AddDialog.CustomRequired'));
      return null;
    }

    return {
      templateId: '',
      name,
      iconPath,
      totalTurns: normalizeInteger(
        form?.elements?.timedEffectDurationTurns?.value,
        DEFAULT_TIMED_EFFECT_DURATION,
        { min: 1, max: 999 },
      ),
      source: lightSourceConfig.isLightSource ? getSourceDisplayLabel({ sourceLabel }, sourceTokenName) : freeformSource,
      sourceLabel: lightSourceConfig.isLightSource ? getSourceDisplayLabel({ sourceLabel }, sourceTokenName) : freeformSource,
      sourceTokenId: lightSourceConfig.isLightSource ? sourceTokenId : '',
      sourceSceneId: lightSourceConfig.isLightSource ? sourceSceneId : '',
      ...lightSourceConfig,
    };
  }

  const template = templates.find((entry) => entry.id === templateId);
  if (!template) return null;

  return {
    templateId: template.id,
    name: template.name,
    iconPath: template.iconPath,
    totalTurns: normalizeInteger(
      form?.elements?.timedEffectDurationTurns?.value,
      template.defaultDurationTurns ?? DEFAULT_TIMED_EFFECT_DURATION,
      { min: 1, max: 999 },
    ),
    source: lightSourceConfig.isLightSource ? getSourceDisplayLabel({ sourceLabel }, sourceTokenName) : freeformSource,
    sourceLabel: lightSourceConfig.isLightSource ? getSourceDisplayLabel({ sourceLabel }, sourceTokenName) : freeformSource,
    sourceTokenId: lightSourceConfig.isLightSource ? sourceTokenId : '',
    sourceSceneId: lightSourceConfig.isLightSource ? sourceSceneId : '',
    ...lightSourceConfig,
  };
}

async function promptForTimedEffect() {
  const templates = getTimedEffectTemplates();
  const sourceOptions = buildTimedEffectSourceOptions();
  const sourceTokens = getCurrentSceneSourceTokens();
  const selectedTokenId = getSingleControlledTokenId();

  const dialogContent = `
    <div class="expedition-tracker-dialog-form expedition-timed-effect-dialog-form">
      <div class="form-group">
        <label>${localizeTracker('TimedEffects.AddDialog.TemplateLabel')}</label>
        <select name="timedEffectTemplateId">
          ${buildTimedEffectTemplateOptions(templates)}
        </select>
      </div>
      <div class="expedition-timed-effect-custom-fields" data-timed-effect-custom-fields hidden>
        <div class="form-group">
          <label>${localizeTracker('TimedEffects.AddDialog.CustomNameLabel')}</label>
          <input type="text" name="timedEffectCustomName" disabled>
        </div>
        <div class="form-group">
          <label>${localizeTracker('TimedEffects.AddDialog.CustomIconLabel')}</label>
          <div class="expedition-timed-effect-icon-path-field">
            <input type="text" name="timedEffectCustomIconPath" disabled>
            <button type="button" data-action="browse-custom-timed-effect-icon">
              ${localizeTracker('TimedEffects.AddDialog.BrowseIconAction')}
            </button>
          </div>
        </div>
      </div>
      <div class="form-group">
        <label>${localizeTracker('TimedEffects.AddDialog.DurationLabel')}</label>
        <input type="number" name="timedEffectDurationTurns" value="${templates[0]?.defaultDurationTurns ?? DEFAULT_TIMED_EFFECT_DURATION}" min="1" step="1" data-is-auto-default="true">
      </div>
      <div class="form-group" data-timed-effect-freeform-source-field>
        <label>${localizeTracker('TimedEffects.AddDialog.SourceLabel')}</label>
        <input type="text" name="timedEffectSource" list="timed-effect-source-options" placeholder="${localizeTracker('TimedEffects.AddDialog.SourcePlaceholder')}">
        <datalist id="timed-effect-source-options">
          ${sourceOptions}
        </datalist>
      </div>
      <div class="form-group">
        <label>
          <input type="checkbox" name="timedEffectIsLightSource" data-is-auto-default="true">
          ${localizeTracker('TimedEffects.AddDialog.LightSourceLabel')}
        </label>
      </div>
      <div class="form-group" data-timed-effect-source-token-field hidden>
        <label>${localizeTracker('TimedEffects.AddDialog.SourceTokenLabel')}</label>
        <select name="timedEffectSourceTokenId" disabled>
          ${buildTimedEffectSourceTokenOptions(sourceTokens, selectedTokenId)}
        </select>
      </div>
      <div class="form-group" data-timed-effect-source-label-field hidden>
        <label>${localizeTracker('TimedEffects.AddDialog.SourceLabelOverride')}</label>
        <input type="text" name="timedEffectSourceLabel" placeholder="${localizeTracker('TimedEffects.AddDialog.SourceLabelOverridePlaceholder')}" disabled>
      </div>
      <div class="expedition-timed-effect-light-fields" data-timed-effect-light-fields hidden>
        <div class="form-group">
          <label>${localizeTracker('TimedEffects.AddDialog.BrightLightLabel')}</label>
          <input type="number" name="timedEffectBrightLight" value="${DEFAULT_LIGHT_SOURCE_CONFIG.brightLight}" min="0" step="1" data-is-auto-default="true" disabled>
        </div>
        <div class="form-group">
          <label>${localizeTracker('TimedEffects.AddDialog.DimLightLabel')}</label>
          <input type="number" name="timedEffectDimLight" value="${DEFAULT_LIGHT_SOURCE_CONFIG.dimLight}" min="0" step="1" data-is-auto-default="true" disabled>
        </div>
        <div class="form-group">
          <label>${localizeTracker('TimedEffects.AddDialog.EmissionAngleLabel')}</label>
          <input type="number" name="timedEffectEmissionAngle" value="${DEFAULT_LIGHT_SOURCE_CONFIG.emissionAngle}" min="1" max="360" step="1" data-is-auto-default="true" disabled>
        </div>
      </div>
    </div>
  `;

  let timedEffectConfig = null;
  await new Promise((resolve) => {
    const dialog = new foundry.applications.api.DialogV2({
      window: {
        title: localizeTracker('TimedEffects.AddDialog.Title'),
      },
      classes: ['expedition-tracker-ui', 'expedition-timed-effect-dialog'],
      content: dialogContent,
      buttons: [
        {
          action: 'add',
          label: localizeTracker('TimedEffects.AddAction'),
          default: true,
          callback: (_event, button) => {
            timedEffectConfig = parseTimedEffectConfig(button.form, templates);
          },
        },
        {
          action: 'cancel',
          label: localizeTracker('Common.Cancel'),
        },
      ],
      submit: () => resolve(),
    });

    dialog.addEventListener('render', () => initializeTimedEffectDialog(dialog, templates));
    dialog.render(true);
  });

  return timedEffectConfig;
}

function getTimedEffectManageDisplayLabelValue(effect) {
  const tokenName = getSceneTokenNameById(effect.sourceTokenId, effect.sourceSceneId);
  return effect.sourceLabel && effect.sourceLabel !== tokenName
    ? effect.sourceLabel
    : '';
}

function buildTimedEffectManageDialogContent(effect) {
  const isCustom = !effect.templateId;
  const isLightSource = isEffectLightSource(effect);
  const tokenName = getSceneTokenNameById(effect.sourceTokenId, effect.sourceSceneId)
    || localizeTracker('TimedEffects.ManageDialog.UnknownToken');

  const nameField = isCustom
    ? `
      <div class="form-group">
        <label>${localizeTracker('TimedEffects.ManageDialog.NameLabel')}</label>
        <input type="text" name="timedEffectName" value="${escapeHtml(effect.name)}">
      </div>
    `
    : `
      <div class="form-group">
        <label>${localizeTracker('TimedEffects.ManageDialog.NameLabel')}</label>
        <input type="text" name="timedEffectName" value="${escapeHtml(effect.name)}" readonly disabled>
      </div>
    `;

  const sourceField = isLightSource
    ? `
      <div class="form-group">
        <label>${localizeTracker('TimedEffects.ManageDialog.TokenLabel')}</label>
        <input type="text" value="${escapeHtml(tokenName)}" readonly disabled>
      </div>
      <div class="form-group">
        <label>${localizeTracker('TimedEffects.ManageDialog.DisplayLabel')}</label>
        <input type="text" name="timedEffectSourceLabel" value="${escapeHtml(getTimedEffectManageDisplayLabelValue(effect))}" placeholder="${escapeHtml(localizeTracker('TimedEffects.ManageDialog.DisplayLabelPlaceholder'))}">
      </div>
    `
    : `
      <div class="form-group">
        <label>${localizeTracker('TimedEffects.ManageDialog.SourceLabel')}</label>
        <input type="text" name="timedEffectSource" value="${escapeHtml(effect.source)}" placeholder="${escapeHtml(localizeTracker('TimedEffects.ManageDialog.SourcePlaceholder'))}">
      </div>
    `;

  const lightFields = isLightSource
    ? `
      <div class="expedition-timed-effect-light-fields">
        <div class="form-group">
          <label>${localizeTracker('TimedEffects.AddDialog.BrightLightLabel')}</label>
          <input type="number" name="timedEffectBrightLight" value="${effect.brightLight}" min="0" step="1">
        </div>
        <div class="form-group">
          <label>${localizeTracker('TimedEffects.AddDialog.DimLightLabel')}</label>
          <input type="number" name="timedEffectDimLight" value="${effect.dimLight}" min="0" step="1">
        </div>
        <div class="form-group">
          <label>${localizeTracker('TimedEffects.AddDialog.EmissionAngleLabel')}</label>
          <input type="number" name="timedEffectEmissionAngle" value="${effect.emissionAngle}" min="1" max="360" step="1">
        </div>
      </div>
    `
    : '';

  return `
    <div class="expedition-tracker-dialog-form expedition-timed-effect-manage-form">
      ${nameField}
      <div class="form-group">
        <label>${localizeTracker('TimedEffects.ManageDialog.RemainingTurns')}</label>
        <input type="number" name="timedEffectRemainingTurns" value="${effect.remainingTurns}" min="1" step="1">
      </div>
      <div class="form-group">
        <label>${localizeTracker('TimedEffects.ManageDialog.TotalTurns')}</label>
        <input type="number" name="timedEffectTotalTurns" value="${effect.totalTurns}" min="1" step="1">
      </div>
      ${sourceField}
      ${lightFields}
    </div>
  `;
}

function parseManagedTimedEffectConfig(form, effect) {
  const totalTurns = normalizeInteger(form?.elements?.timedEffectTotalTurns?.value, effect.totalTurns, { min: 1, max: 999 });
  const remainingTurns = normalizeInteger(form?.elements?.timedEffectRemainingTurns?.value, effect.remainingTurns, { min: 1, max: 999 });
  const normalizedRemainingTurns = remainingTurns;
  const normalizedTotalTurns = Math.max(totalTurns, normalizedRemainingTurns);

  const nextConfig = {
    ...effect,
    totalTurns: normalizedTotalTurns,
    remainingTurns: normalizedRemainingTurns,
  };

  if (!effect.templateId) {
    const customName = String(form?.elements?.timedEffectName?.value ?? '').trim();
    if (!customName) {
      ui.notifications.warn(localizeTracker('TimedEffects.ManageDialog.CustomNameRequired'));
      return null;
    }
    nextConfig.name = customName;
  }

  if (isEffectLightSource(effect)) {
    const tokenName = getSceneTokenNameById(effect.sourceTokenId, effect.sourceSceneId);
    const sourceLabel = String(form?.elements?.timedEffectSourceLabel?.value ?? '').trim();
    nextConfig.sourceLabel = sourceLabel;
    nextConfig.source = getSourceDisplayLabel({ sourceLabel }, tokenName);
    nextConfig.brightLight = normalizeInteger(form?.elements?.timedEffectBrightLight?.value, effect.brightLight, { min: 0, max: 999 });
    nextConfig.dimLight = normalizeInteger(form?.elements?.timedEffectDimLight?.value, effect.dimLight, { min: 0, max: 999 });
    nextConfig.emissionAngle = normalizeInteger(form?.elements?.timedEffectEmissionAngle?.value, effect.emissionAngle, { min: 1, max: 360 });
  } else {
    const source = String(form?.elements?.timedEffectSource?.value ?? '').trim();
    nextConfig.source = source;
    nextConfig.sourceLabel = source;
  }

  return nextConfig;
}

async function promptForTimedEffectManagement(effect) {
  let result = null;

  await new Promise((resolve) => {
    new foundry.applications.api.DialogV2({
      window: {
        title: localizeTracker('TimedEffects.ManageDialog.Title', { name: effect.name }),
      },
      classes: ['expedition-tracker-ui', 'expedition-timed-effect-dialog'],
      content: buildTimedEffectManageDialogContent(effect),
      buttons: [
        {
          action: 'save',
          label: localizeTracker('TimedEffects.ManageDialog.SaveAction'),
          default: true,
          callback: (_event, button) => {
            result = {
              action: 'save',
              config: parseManagedTimedEffectConfig(button.form, effect),
            };
          },
        },
        {
          action: effect.isPaused ? 'resume' : 'pause',
          label: effect.isPaused
            ? localizeTracker('TimedEffects.ManageDialog.ResumeAction')
            : localizeTracker('TimedEffects.ManageDialog.PauseAction'),
          callback: (_event, button) => {
            result = {
              action: effect.isPaused ? 'resume' : 'pause',
              config: parseManagedTimedEffectConfig(button.form, effect),
            };
          },
        },
        {
          action: 'remove',
          label: localizeTracker('TimedEffects.RemoveAction'),
          callback: () => {
            result = {
              action: 'remove',
            };
          },
        },
        {
          action: 'cancel',
          label: localizeTracker('Common.Cancel'),
        },
      ],
      submit: () => resolve(),
    }).render(true);
  });

  if (result?.config === null) return null;
  return result;
}

async function addTimedEffect() {
  const state = getTrackerState();
  if (!state.isActive || !isTimedEffectTrackEnabled()) return;

  const timedEffectConfig = await promptForTimedEffect();
  if (!timedEffectConfig) return;

  const timedEffect = normalizeTimedEffectInstance({
    id: foundry.utils.randomID(),
    ...timedEffectConfig,
    remainingTurns: timedEffectConfig.totalTurns,
  });
  if (!timedEffect) return;
  const logEntry = createTimedEffectActivatedEntry(timedEffect);

  await queueTrackerStateUpdate(async (currentState) => {
    const nextState = {
      ...currentState,
      timedEffects: [
        ...currentState.timedEffects,
        timedEffect,
      ],
      logEntries: appendLogEntry(currentState.logEntries, logEntry),
    };

    if (nextState.logToJournal && nextState.journalEntryId && nextState.journalPageId) {
      await appendToExpeditionJournalPage({
        journalEntryId: nextState.journalEntryId,
        journalPageId: nextState.journalPageId,
        entry: logEntry.text,
      });
    }

    return nextState;
  });

  if (isEffectLightSource(timedEffect)) {
    await recalculateTokenLightForToken(getTrackerState(), timedEffect.sourceTokenId, timedEffect.sourceSceneId);
  }
}

async function manageTimedEffect(effectId) {
  const currentEffect = getTimedEffectById(effectId);
  if (!currentEffect) return;

  const result = await promptForTimedEffectManagement(currentEffect);
  if (!result) return;

  if (result.action === 'remove') {
    await removeTimedEffect(effectId);
    return;
  }

  const nextTimedEffect = normalizeTimedEffectInstance({
    ...currentEffect,
    ...result.config,
    isPaused: result.action === 'pause'
      ? true
      : result.action === 'resume'
        ? false
        : currentEffect.isPaused,
  });
  if (!nextTimedEffect) return;

  const logEntry = result.action === 'pause'
    ? createTimedEffectPausedEntry(nextTimedEffect)
    : result.action === 'resume'
      ? createTimedEffectResumedEntry(nextTimedEffect)
      : null;

  await queueTrackerStateUpdate(async (state) => {
    const existingEffect = state.timedEffects.find((effect) => effect.id === effectId);
    if (!existingEffect) return state;

    let nextLogEntries = state.logEntries;
    if (logEntry) {
      nextLogEntries = appendLogEntry(nextLogEntries, logEntry);
    }

    const nextState = {
      ...state,
      logEntries: nextLogEntries,
      timedEffects: state.timedEffects.map((effect) => (effect.id === effectId ? nextTimedEffect : effect)),
    };

    if (logEntry && nextState.logToJournal && nextState.journalEntryId && nextState.journalPageId) {
      await appendToExpeditionJournalPage({
        journalEntryId: nextState.journalEntryId,
        journalPageId: nextState.journalPageId,
        entry: logEntry.text,
      });
    }

    return nextState;
  });

  if (isEffectLightSource(nextTimedEffect)) {
    await recalculateTokenLightForToken(getTrackerState(), nextTimedEffect.sourceTokenId, nextTimedEffect.sourceSceneId);
  }
}

async function removeTimedEffect(effectId) {
  const currentState = getTrackerState();
  const timedEffect = currentState.timedEffects.find((effect) => effect.id === effectId);
  if (!timedEffect) return;

  if (isEffectLightSource(timedEffect)) {
    const removalChoice = await promptForTimedEffectRemoval(timedEffect);
    if (removalChoice === 'cancel') return;

    await queueTrackerStateUpdate((state) => ({
      ...state,
      timedEffects: state.timedEffects.filter((effect) => effect.id !== effectId),
    }));

    if (removalChoice === 'remove-and-recalculate') {
      await recalculateTokenLightForToken(getTrackerState(), timedEffect.sourceTokenId, timedEffect.sourceSceneId);
    }
    return;
  }

  const confirmed = await confirmExpeditionAction({
    title: localizeTracker('TimedEffects.RemoveDialog.Title'),
    body: localizeTracker('TimedEffects.RemoveDialog.Body', { name: timedEffect.name }),
    confirmLabel: localizeTracker('TimedEffects.RemoveAction'),
  });
  if (!confirmed) return;

  await queueTrackerStateUpdate((state) => ({
    ...state,
    timedEffects: state.timedEffects.filter((effect) => effect.id !== effectId),
  }));
}

function getTimedEffectTooltipElement() {
  let tooltip = document.getElementById(TIMED_EFFECT_TOOLTIP_ID);
  if (tooltip) return tooltip;

  tooltip = document.createElement('div');
  tooltip.id = TIMED_EFFECT_TOOLTIP_ID;
  tooltip.className = 'expedition-timed-effect-tooltip';
  tooltip.hidden = true;
  document.body.append(tooltip);
  return tooltip;
}

function hideTimedEffectTooltip() {
  const tooltip = document.getElementById(TIMED_EFFECT_TOOLTIP_ID);
  if (!tooltip) return;
  tooltip.hidden = true;
  tooltip.replaceChildren();
}

function positionTimedEffectTooltip(tooltip, anchor, clientX, clientY) {
  const anchorRect = anchor.getBoundingClientRect();
  const left = Number.isFinite(clientX) ? clientX + 12 : anchorRect.left + (anchorRect.width / 2);
  const top = Number.isFinite(clientY) ? clientY + 12 : anchorRect.bottom + 8;
  const maxLeft = window.innerWidth - tooltip.offsetWidth - 12;
  const maxTop = window.innerHeight - tooltip.offsetHeight - 12;

  tooltip.style.left = `${Math.max(12, Math.min(left, maxLeft))}px`;
  tooltip.style.top = `${Math.max(12, Math.min(top, maxTop))}px`;
}

function showTimedEffectTooltip(anchor, effectId, event = null) {
  const effect = getTimedEffectById(effectId);
  if (!effect) return;

  const tooltip = getTimedEffectTooltipElement();
  tooltip.innerHTML = buildTimedEffectTooltipHtml(effect);
  tooltip.hidden = false;
  positionTimedEffectTooltip(tooltip, anchor, event?.clientX, event?.clientY);
}

async function promptForTimedEffectRemoval(timedEffect) {
  let action = 'cancel';

  await new Promise((resolve) => {
    new foundry.applications.api.DialogV2({
      window: {
        title: localizeTracker('TimedEffects.RemoveDialog.Title'),
      },
      classes: ['expedition-tracker-ui'],
      content: `<p>${escapeHtml(localizeTracker('TimedEffects.RemoveDialog.LightSourceBody', { name: timedEffect.name }))}</p>`,
      buttons: [
        {
          action: 'remove-only',
          label: localizeTracker('TimedEffects.RemoveDialog.RemoveOnlyAction'),
          callback: () => {
            action = 'remove-only';
          },
        },
        {
          action: 'remove-and-recalculate',
          label: localizeTracker('TimedEffects.RemoveDialog.RemoveAndRecalculateAction'),
          default: true,
          callback: () => {
            action = 'remove-and-recalculate';
          },
        },
        {
          action: 'cancel',
          label: localizeTracker('Common.Cancel'),
        },
      ],
      submit: () => resolve(),
    }).render(true);
  });

  return action;
}

async function confirmExpeditionAction({
  title,
  body,
  confirmLabel,
}) {
  let confirmed = false;

  await new Promise((resolve) => {
    new foundry.applications.api.DialogV2({
      window: {
        title,
      },
      classes: ['expedition-tracker-ui'],
      content: `<p>${body}</p>`,
      buttons: [
        {
          action: 'confirm',
          label: confirmLabel,
          default: true,
          callback: () => {
            confirmed = true;
          },
        },
        {
          action: 'cancel',
          label: localizeTracker('Common.Cancel'),
        },
      ],
      submit: () => resolve(),
    }).render(true);
  });

  return confirmed;
}

async function startNewExpedition() {
  const currentState = getTrackerState();
  const initialTimekeeping = normalizeTimekeepingConfig(currentState.timekeeping);
  const initialTimekeepingSummary = getTimekeepingSummaryLines(initialTimekeeping);
  const initialWanderingCheck = normalizeWanderingCheckConfig(currentState.wanderingCheck);
  if (currentState.isActive) {
    const confirmed = await confirmExpeditionAction({
      title: localizeTracker('Expedition.StartConfirmTitle'),
      body: localizeTracker('Expedition.StartConfirmBody'),
      confirmLabel: localizeTracker('Expedition.StartConfirmYes'),
    });
    if (!confirmed) return;
  }

  const dialogContent = `
    <div class="expedition-tracker-dialog-form">
      <div class="form-group">
        <label>${localizeTracker('Expedition.NameLabel')}</label>
        <input type="text" name="expeditionName" placeholder="${localizeTracker('Expedition.NamePlaceholder')}" autofocus>
      </div>
      <div class="form-group">
        <label>
          <input type="checkbox" name="logToJournal">
          ${localizeTracker('Journal.LogLabel')}
        </label>
        <p class="hint">${localizeTracker('Journal.LogHint')}</p>
      </div>
      <fieldset class="form-group stacked expedition-timekeeping-fieldset">
        <legend>${localizeTracker('Time.SettingsLegend')}</legend>
        <div class="form-fields">
          <label>${localizeTracker('Time.ProcedureLabel')}</label>
          <select name="timekeepingPreset">
            <option value="dungeon"${initialTimekeeping.preset === 'dungeon' ? ' selected' : ''}>${localizeTracker('Time.Presets.Dungeon')}</option>
            <option value="overland"${initialTimekeeping.preset === 'overland' ? ' selected' : ''}>${localizeTracker('Time.Presets.Overland')}</option>
            <option value="custom"${initialTimekeeping.preset === 'custom' ? ' selected' : ''}>${localizeTracker('Time.Presets.Custom')}</option>
          </select>
        </div>
        <div class="expedition-timekeeping-summary" data-timekeeping-summary${initialTimekeeping.preset === 'custom' ? ' hidden' : ''}>
          <div class="expedition-timekeeping-summary-line" data-summary-line="increment">${initialTimekeepingSummary.increment}</div>
          <div class="expedition-timekeeping-summary-line" data-summary-line="cycle">${initialTimekeepingSummary.cycle}</div>
          <div class="expedition-timekeeping-summary-line" data-summary-line="cadence">${initialTimekeepingSummary.cadence}</div>
        </div>
        <div class="expedition-timekeeping-custom-fields" data-timekeeping-custom-fields${initialTimekeeping.preset === 'custom' ? '' : ' hidden'}>
          <div class="form-fields">
            <label>${localizeTracker('Time.IncrementLabelSingular')}</label>
            <input type="text" name="timekeepingIncrementLabelSingular" value="${escapeHtml(initialTimekeeping.incrementLabelSingular)}"${initialTimekeeping.preset === 'custom' ? '' : ' disabled'}>
          </div>
          <div class="form-fields">
            <label>${localizeTracker('Time.IncrementLabelPlural')}</label>
            <input type="text" name="timekeepingIncrementLabelPlural" value="${escapeHtml(initialTimekeeping.incrementLabelPlural)}"${initialTimekeeping.preset === 'custom' ? '' : ' disabled'}>
          </div>
          <div class="form-fields">
            <label>${localizeTracker('Time.CycleLabelSingular')}</label>
            <input type="text" name="timekeepingCycleLabelSingular" value="${escapeHtml(initialTimekeeping.cycleLabelSingular)}"${initialTimekeeping.preset === 'custom' ? '' : ' disabled'}>
          </div>
          <div class="form-fields">
            <label>${localizeTracker('Time.CycleLabelPlural')}</label>
            <input type="text" name="timekeepingCycleLabelPlural" value="${escapeHtml(initialTimekeeping.cycleLabelPlural)}"${initialTimekeeping.preset === 'custom' ? '' : ' disabled'}>
          </div>
          <div class="form-fields">
            <label>${localizeTracker('Time.IncrementsPerCycle')}</label>
            <input type="number" name="timekeepingIncrementsPerCycle" value="${initialTimekeeping.incrementsPerCycle}" min="1" step="1"${initialTimekeeping.preset === 'custom' ? '' : ' disabled'}>
          </div>
        </div>
      </fieldset>
      <fieldset class="form-group stacked expedition-wandering-fieldset">
        <legend>${localizeTracker('Wandering.SettingsLegend')}</legend>
        <div class="form-fields">
          <label>${localizeTracker('Wandering.EncounterChanceLabel')}</label>
          <div class="form-fields">
            <input type="number" name="wanderingEncounterThreshold" value="${initialWanderingCheck.encounterThreshold}" min="1" step="1">
            <span>${localizeTracker('Wandering.ChanceSeparator')}</span>
            <input type="number" name="wanderingDieFaces" value="${initialWanderingCheck.dieFaces}" min="2" step="1">
          </div>
        </div>
        <div class="form-fields">
          <label>${localizeTracker('Wandering.EveryLabel')}</label>
          <input type="number" name="wanderingCheckEveryTurns" value="${initialWanderingCheck.checkEveryTurns}" min="1" step="1">
        </div>
        <div class="form-group">
          <label>
            <input type="checkbox" name="wanderingAutoCheck"${initialWanderingCheck.autoCheck ? ' checked' : ''}>
            ${localizeTracker('Wandering.AutoCheckLabel')}
          </label>
          <p class="hint">${localizeTracker('Wandering.AutoCheckHint')}</p>
        </div>
        <p class="hint">${localizeTracker('Wandering.SettingsHint')}</p>
      </fieldset>
    </div>
  `;

  let expeditionConfig = null;
  await new Promise((resolve) => {
    const dialog = new foundry.applications.api.DialogV2({
      position: {
        width: 700,
      },
      window: {
        title: localizeTracker('Expedition.StartDialogTitle'),
      },
      classes: ['expedition-tracker-ui', 'expedition-config-dialog'],
      content: dialogContent,
      buttons: [
        {
          action: 'start',
          label: localizeTracker('Expedition.StartAction'),
          default: true,
          callback: (_event, button) => {
            expeditionConfig = parseExpeditionConfig(button.form);
          },
        },
        {
          action: 'cancel',
          label: localizeTracker('Common.Cancel'),
        },
      ],
      submit: () => resolve(),
    });
    dialog.addEventListener('render', () => initializeExpeditionDialog(dialog));
    dialog.render({ force: true });
  });

  if (!expeditionConfig) return;

  if (expeditionConfig.customTimekeepingIncomplete) {
    ui.notifications.warn(localizeTracker('Time.CustomRequired'));
    return;
  }

  if (expeditionConfig.logToJournal && !expeditionConfig.expeditionName) {
    ui.notifications.warn(localizeTracker('Journal.NameRequired'));
    return;
  }

  let journalData = {
    journalEntryId: '',
    journalPageId: '',
  };
  if (expeditionConfig.logToJournal) {
    try {
      journalData = await createExpeditionJournalPage(expeditionConfig.expeditionName, expeditionConfig.timekeeping);
    } catch (error) {
      console.error(`${MODULE_ID} | Failed to create expedition journal page`, error);
      ui.notifications.error(String(error?.message ?? localizeTracker('Journal.Unavailable')));
      return;
    }
  }

  if (currentState.isActive && currentState.logToJournal && currentState.journalEntryId && currentState.journalPageId) {
    try {
      await appendToExpeditionJournalPage({
        journalEntryId: currentState.journalEntryId,
        journalPageId: currentState.journalPageId,
        entry: localizeTracker('Journal.Ended'),
      });
    } catch (error) {
      console.error(`${MODULE_ID} | Failed to append prior expedition end to journal`, error);
      ui.notifications.error(String(error?.message ?? localizeTracker('Journal.Unavailable')));
      return;
    }
  }

  await removeActiveExpeditionLights(currentState.timedEffects);
  await game.settings.set(MODULE_ID, TRACKER_SETTING_KEY, {
    ...getDefaultTrackerState(),
    isActive: true,
    expeditionName: expeditionConfig.expeditionName,
    logToJournal: expeditionConfig.logToJournal,
    timekeeping: expeditionConfig.timekeeping,
    wanderingCheck: expeditionConfig.wanderingCheck,
    ...journalData,
  });
}

async function endExpedition() {
  const currentState = getTrackerState();
  if (!currentState.isActive) return;

  const confirmed = await confirmExpeditionAction({
    title: localizeTracker('Expedition.EndConfirmTitle'),
    body: localizeTracker('Expedition.EndConfirmBody'),
    confirmLabel: localizeTracker('Expedition.EndConfirmYes'),
  });
  if (!confirmed) return;

  if (currentState.logToJournal && currentState.journalEntryId && currentState.journalPageId) {
    try {
      await appendToExpeditionJournalPage({
        journalEntryId: currentState.journalEntryId,
        journalPageId: currentState.journalPageId,
        entry: localizeTracker('Journal.Ended'),
      });
    } catch (error) {
      console.error(`${MODULE_ID} | Failed to append expedition end to journal`, error);
      ui.notifications.error(String(error?.message ?? localizeTracker('Journal.Unavailable')));
    }
  }

  await removeActiveExpeditionLights(currentState.timedEffects);
  await game.settings.set(MODULE_ID, TRACKER_SETTING_KEY, getDefaultTrackerState());
}

async function handleTurnAction(action) {
  if (action === 'other') {
    const customLabel = await promptForOtherTurnLabel();
    if (!customLabel) {
      ui.notifications.warn(localizeTracker('OtherTurn.Empty'));
      return;
    }
    await recordTurnAction(customLabel);
    return;
  }

  await recordTurnAction(getTurnActionLabel(action));
}

async function handleCurrentTurnNote() {
  const state = getTrackerState();
  if (!state.isActive || state.turnCount <= 0) {
    ui.notifications.warn(localizeTracker('Notes.Unavailable'));
    return;
  }

  const note = await promptForCurrentTurnNote();
  if (!note) {
    ui.notifications.warn(localizeTracker('Notes.Empty'));
    return;
  }

  await recordCurrentTurnNote(note);
}

function bindTrackerOverlayEvents(overlay) {
  hideTimedEffectTooltip();

  overlay.querySelector('.timed-effects-add-action')?.addEventListener('click', async (event) => {
    event.preventDefault();
    await addTimedEffect();
  });

  overlay.querySelectorAll('[data-timed-effect-manage]').forEach((button) => {
    button.addEventListener('click', async (event) => {
      event.preventDefault();
      const effectId = String(event.currentTarget.dataset.timedEffectManage ?? '').trim();
      if (!effectId) return;
      hideTimedEffectTooltip();
      await manageTimedEffect(effectId);
    });

    button.addEventListener('mouseenter', (event) => {
      const effectId = String(event.currentTarget.dataset.timedEffectManage ?? '').trim();
      if (!effectId) return;
      showTimedEffectTooltip(event.currentTarget, effectId, event);
    });

    button.addEventListener('mousemove', (event) => {
      const tooltip = document.getElementById(TIMED_EFFECT_TOOLTIP_ID);
      if (!tooltip || tooltip.hidden) return;
      positionTimedEffectTooltip(tooltip, event.currentTarget, event.clientX, event.clientY);
    });

    button.addEventListener('mouseleave', () => {
      hideTimedEffectTooltip();
    });
  });

  overlay.querySelector('.referee-log-toggle')?.addEventListener('click', async (event) => {
    event.preventDefault();
    isLogCollapsed = !isLogCollapsed;
    await ExpeditionTrackerHud.sync();
  });

  overlay.querySelector('.wandering-check-action')?.addEventListener('click', async (event) => {
    event.preventDefault();
    await runWanderingCheck();
  });

  overlay.querySelector('.other-turn-action')?.addEventListener('click', async (event) => {
    event.preventDefault();
    await handleTurnAction('other');
  });

  overlay.querySelector('.note-current-turn-action')?.addEventListener('click', async (event) => {
    event.preventDefault();
    await handleCurrentTurnNote();
  });

  overlay.querySelector('.reset-expedition-action')?.addEventListener('click', async (event) => {
    event.preventDefault();
    await startNewExpedition();
  });

  overlay.querySelector('.end-expedition-action')?.addEventListener('click', async (event) => {
    event.preventDefault();
    await endExpedition();
  });

  overlay.querySelectorAll('[data-turn-action]').forEach((button) => {
    button.addEventListener('click', async (event) => {
      event.preventDefault();
      const action = String(event.currentTarget.dataset.turnAction ?? '').trim();
      if (!action) return;
      await handleTurnAction(action);
    });
  });
}

function getHotbarReferenceRect() {
  const actionBar = document.querySelector('#hotbar #action-bar');
  const hotbar = document.getElementById('hotbar');
  const reference = actionBar ?? hotbar;
  if (!reference) return null;
  const rect = reference.getBoundingClientRect();
  return rect.width > 0 ? rect : null;
}

function normalizeTrackerOverlayPosition(position = {}) {
  const left = Number(position.left);
  const top = Number(position.top);
  const width = Number(position.width);

  return {
    left: Number.isFinite(left) ? left : null,
    top: Number.isFinite(top) ? top : null,
    width: Number.isFinite(width) ? width : null,
  };
}

function getStoredTrackerOverlayPosition() {
  return normalizeTrackerOverlayPosition(game.settings.get(MODULE_ID, TRACKER_POSITION_SETTING_KEY));
}

function getDefaultTrackerOverlayPosition(overlayHeight = 240) {
  const hotbarRect = getHotbarReferenceRect();
  if (!hotbarRect) return null;

  return {
    left: Math.max(8, hotbarRect.left),
    top: Math.max(8, hotbarRect.top - 8 - overlayHeight),
    width: Math.max(320, hotbarRect.width),
  };
}

function clampTrackerOverlayPosition(position, overlayWidth = 320, overlayHeight = 0) {
  const margin = 8;
  const width = Math.max(320, Math.min(position.width ?? overlayWidth, window.innerWidth - (margin * 2)));
  const maxLeft = Math.max(margin, window.innerWidth - width - margin);
  const maxTop = Math.max(margin, window.innerHeight - overlayHeight - margin);

  return {
    left: Math.min(Math.max(position.left ?? margin, margin), maxLeft),
    top: Math.min(Math.max(position.top ?? margin, margin), maxTop),
    width,
  };
}

async function saveTrackerOverlayPosition(position) {
  const normalized = clampTrackerOverlayPosition(position, position.width, position.height ?? 0);
  await game.settings.set(MODULE_ID, TRACKER_POSITION_SETTING_KEY, {
    left: Math.round(normalized.left),
    top: Math.round(normalized.top),
    width: Math.round(normalized.width),
  });
}

function positionTrackerOverlay(overlay) {
  overlay.style.removeProperty('left');
  overlay.style.removeProperty('top');
  overlay.style.removeProperty('width');
  overlay.style.removeProperty('bottom');

  const defaultPosition = getDefaultTrackerOverlayPosition(overlay.offsetHeight ?? 240);
  const storedPosition = getStoredTrackerOverlayPosition();
  const requestedPosition = storedPosition.left !== null && storedPosition.top !== null
    ? {
      left: storedPosition.left,
      top: storedPosition.top,
      width: storedPosition.width ?? defaultPosition?.width ?? overlay.offsetWidth ?? 320,
    }
    : defaultPosition;

  if (!requestedPosition) return;

  const clampedPosition = clampTrackerOverlayPosition(
    requestedPosition,
    requestedPosition.width ?? overlay.offsetWidth ?? 320,
    overlay.offsetHeight ?? 0,
  );

  overlay.style.left = `${clampedPosition.left}px`;
  overlay.style.top = `${clampedPosition.top}px`;
  overlay.style.width = `${clampedPosition.width}px`;
}

function endTrackerOverlayDrag() {
  if (!trackerDragState) return;
  window.removeEventListener('mousemove', trackerDragState.onMouseMove);
  window.removeEventListener('mouseup', trackerDragState.onMouseUp);
  trackerDragState.overlay.classList.remove('expedition-tracker-dragging');
  trackerDragState = null;
}

function initializeTrackerOverlayDrag(overlay) {
  if (!overlay) return;
  const dragHandle = overlay.querySelector('.expedition-tracker-header');
  if (!dragHandle || dragHandle.dataset.dragInitialized === 'true') return;
  dragHandle.dataset.dragInitialized = 'true';

  dragHandle.classList.add('expedition-tracker-drag-handle');
  dragHandle.addEventListener('mousedown', (event) => {
    if (event.button !== 0) return;
    if (event.target.closest('button, input, select, textarea, a')) return;

    event.preventDefault();
    endTrackerOverlayDrag();

    const overlayRect = overlay.getBoundingClientRect();
    const initialOffsetX = event.clientX - overlayRect.left;
    const initialOffsetY = event.clientY - overlayRect.top;

    const onMouseMove = (moveEvent) => {
      const nextPosition = clampTrackerOverlayPosition({
        left: moveEvent.clientX - initialOffsetX,
        top: moveEvent.clientY - initialOffsetY,
        width: overlayRect.width,
      }, overlayRect.width, overlayRect.height);

      overlay.style.left = `${nextPosition.left}px`;
      overlay.style.top = `${nextPosition.top}px`;
      overlay.style.width = `${nextPosition.width}px`;
      overlay.style.removeProperty('bottom');
    };

    const onMouseUp = async () => {
      const finalRect = overlay.getBoundingClientRect();
      endTrackerOverlayDrag();
      await saveTrackerOverlayPosition({
        left: finalRect.left,
        top: finalRect.top,
        width: finalRect.width,
        height: finalRect.height,
      });
    };

    trackerDragState = {
      overlay,
      onMouseMove,
      onMouseUp,
    };

    overlay.classList.add('expedition-tracker-dragging');
    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  });
}

export class ExpeditionTrackerHud {
  static async sync() {
    const existing = document.getElementById(TRACKER_OVERLAY_ID);

    if (!canvas?.ready || !isExpeditionTrackerEnabled()) {
      this.remove();
      return;
    }

    const markup = await buildTrackerMarkup();
    const overlay = document.getElementById(TRACKER_OVERLAY_ID) ?? existing ?? document.createElement('div');
    overlay.id = TRACKER_OVERLAY_ID;
    overlay.className = 'expedition-tracker-overlay';
    overlay.innerHTML = markup;
    if (!overlay.isConnected) document.body.append(overlay);
    bindTrackerOverlayEvents(overlay);
    initializeTrackerOverlayDrag(overlay);
    positionTrackerOverlay(overlay);
  }

  static remove() {
    endTrackerOverlayDrag();
    hideTimedEffectTooltip();
    document.getElementById(TRACKER_OVERLAY_ID)?.remove();
  }
}
