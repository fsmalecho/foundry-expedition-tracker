export const MODULE_ID = 'expedition-tracker';
export const EXPEDITION_TRACKER_UPDATE_HOOK = 'expeditionTrackerUpdated';
export const TRACKER_ENABLED_SETTING_KEY = 'enableExpeditionTracker';
export const TIMED_EFFECTS_ENABLED_SETTING_KEY = 'enableTimedEffectsTurnTrack';
export const TRACKER_SETTING_KEY = 'expeditionTrackerState';
export const TIMED_EFFECT_TEMPLATES_SETTING_KEY = 'timedEffectTemplates';
export const TRACKER_POSITION_SETTING_KEY = 'expeditionTrackerPosition';
const TIMED_EFFECT_TEMPLATE_MENU_KEY = 'timedEffectTemplatesMenu';
const TRACKER_POSITION_RESET_MENU_KEY = 'trackerPositionResetMenu';
const TIMED_EFFECT_TEMPLATE_CONFIG_TEMPLATE = 'modules/expedition-tracker/templates/timed-effect-templates-config.hbs';
const TRACKER_POSITION_RESET_TEMPLATE = 'modules/expedition-tracker/templates/tracker-position-reset.hbs';
const TIMED_EFFECT_TEMPLATE_CONFIG_MAX_HEIGHT = 600;

function localize(key, data = {}) {
  return data && Object.keys(data).length
    ? game.i18n.format(key, data)
    : game.i18n.localize(key);
}

function localizeTracker(key, data = {}) {
  return localize(`EXPEDITION_TRACKER.${key}`, data);
}

function normalizeInteger(value, fallback, { min = 1, max = Number.MAX_SAFE_INTEGER, allowBlank = false } = {}) {
  const normalizedValue = String(value ?? '').trim();
  if (!normalizedValue) return allowBlank ? null : fallback;

  const parsed = Math.trunc(Number(normalizedValue));
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(max, Math.max(min, parsed));
}

function normalizeTimedEffectTemplateEntry(template = {}, index = 0) {
  const id = String(template.id ?? '').trim() || foundry.utils.randomID();
  const name = String(template.name ?? '').trim();
  const iconPath = String(template.iconPath ?? '').trim();
  const defaultDurationTurns = normalizeInteger(template.defaultDurationTurns, 1, { min: 1, max: 999, allowBlank: true });
  const isLightSource = Boolean(template.isLightSource);
  const brightLight = normalizeInteger(template.brightLight, 0, { min: 0, max: 999, allowBlank: true }) ?? 0;
  const dimLight = normalizeInteger(template.dimLight, 0, { min: 0, max: 999, allowBlank: true }) ?? 0;
  const emissionAngle = normalizeInteger(template.emissionAngle, 360, { min: 1, max: 360, allowBlank: true }) ?? 360;
  const hasAnyValue = Boolean(
    name
    || iconPath
    || String(template.defaultDurationTurns ?? '').trim()
    || isLightSource
    || String(template.brightLight ?? '').trim()
    || String(template.dimLight ?? '').trim()
    || String(template.emissionAngle ?? '').trim(),
  );

  if (!hasAnyValue) return null;
  if (!name || !iconPath) {
    throw new Error(localizeTracker('TimedEffects.TemplateValidation'));
  }

  return {
    id: id || `timed-effect-template-${index + 1}`,
    name,
    iconPath,
    defaultDurationTurns,
    isLightSource,
    brightLight,
    dimLight,
    emissionAngle,
  };
}

export function getDefaultTrackerState() {
  return {
    isActive: false,
    expeditionName: '',
    logToJournal: false,
    journalEntryId: '',
    journalPageId: '',
    turnCount: 0,
    lastTurnLabel: '',
    timekeeping: {
      preset: 'dungeon',
      incrementLabelSingular: 'Turn',
      incrementLabelPlural: 'Turns',
      cycleLabelSingular: 'Hour',
      cycleLabelPlural: 'Hours',
      incrementsPerCycle: 6,
    },
    wanderingCheck: {
      encounterThreshold: 1,
      dieFaces: 6,
      checkEveryTurns: 1,
      autoCheck: false,
    },
    lastWanderingCheck: null,
    logEntries: [],
    timedEffects: [],
  };
}

export function getTimedEffectTemplates() {
  const storedTemplates = game.settings.get(MODULE_ID, TIMED_EFFECT_TEMPLATES_SETTING_KEY);
  if (!Array.isArray(storedTemplates)) return [];

  const templates = [];
  for (const [index, template] of storedTemplates.entries()) {
    try {
      const normalizedTemplate = normalizeTimedEffectTemplateEntry(template, index);
      if (normalizedTemplate) templates.push(normalizedTemplate);
    } catch (error) {
      console.warn(`${MODULE_ID} | Skipping invalid timed effect template`, error);
    }
  }

  return templates;
}

class TimedEffectTemplatesConfig extends FormApplication {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      id: `${MODULE_ID}-timed-effect-templates-config`,
      template: TIMED_EFFECT_TEMPLATE_CONFIG_TEMPLATE,
      title: localizeTracker('TimedEffects.Settings.MenuLabel'),
      width: 860,
      height: 'auto',
      closeOnSubmit: true,
      submitOnClose: false,
      classes: ['expedition-timed-effect-template-config'],
    });
  }

  getData() {
    const templates = getTimedEffectTemplates();
    return {
      templates: templates.length
        ? templates
        : [{
          id: foundry.utils.randomID(),
          name: '',
          iconPath: '',
          defaultDurationTurns: 6,
          isLightSource: false,
          brightLight: 0,
          dimLight: 0,
          emissionAngle: 360,
        }],
    };
  }

  activateListeners(html) {
    super.activateListeners(html);
    this.#syncWindowHeight(html[0]);
    this.#syncAllLightSourceRows(html[0]);

    html[0].addEventListener('click', (event) => {
      const actionButton = event.target.closest('[data-action]');
      if (!actionButton) return;

      const action = String(actionButton.dataset.action ?? '').trim();
      if (action === 'add-template') {
        event.preventDefault();
        this.#appendTemplateRow(html[0].querySelector('[data-template-rows]'));
        this.#syncWindowHeight(html[0]);
        this.#syncAllLightSourceRows(html[0]);
      }

      if (action === 'remove-template') {
        event.preventDefault();
        actionButton.closest('[data-template-row]')?.remove();
        this.#reindexTemplateRows(html[0].querySelector('[data-template-rows]'));
        this.#syncWindowHeight(html[0]);
      }

      if (action === 'browse-icon-path') {
        event.preventDefault();
        const row = actionButton.closest('[data-template-row]');
        const input = row?.querySelector('input[name$=".iconPath"]');
        if (!input) return;

        new FilePicker({
          type: 'imagevideo',
          current: String(input.value ?? '').trim(),
          callback: (path) => {
            input.value = path;
          },
        }).browse();
      }
    });

    html[0].addEventListener('change', (event) => {
      const checkbox = event.target.closest('input[data-light-source-toggle]');
      if (!checkbox) return;
      const row = checkbox.closest('[data-template-row]');
      this.#syncLightSourceRow(row);
      this.#syncWindowHeight(html[0]);
    });
  }

  async _updateObject(_event, formData) {
    const expanded = foundry.utils.expandObject(formData);
    const rawTemplates = Object.values(expanded.templates ?? {});
    const normalizedTemplates = rawTemplates
      .map((template, index) => normalizeTimedEffectTemplateEntry(template, index))
      .filter(Boolean);

    await game.settings.set(MODULE_ID, TIMED_EFFECT_TEMPLATES_SETTING_KEY, normalizedTemplates);
  }

  #appendTemplateRow(container) {
    if (!container) return;
    const row = document.createElement('tr');
    row.dataset.templateRow = 'true';
    row.innerHTML = `
      <td>
        <input type="hidden" value="${foundry.utils.randomID()}">
        <input type="text" value="">
      </td>
      <td>
        <div class="expedition-timed-effect-icon-path-field">
          <input type="text" value="">
          <button type="button" data-action="browse-icon-path">
            ${localizeTracker('TimedEffects.Settings.BrowseIconPath')}
          </button>
        </div>
      </td>
      <td>
        <div class="expedition-timed-effect-template-config-stack">
          <label class="expedition-timed-effect-inline-field">
            <span>${localizeTracker('TimedEffects.Settings.TemplateDefaultDuration')}</span>
            <input type="number" value="6" min="1" step="1">
          </label>
          <label class="expedition-timed-effect-inline-checkbox">
            <input type="checkbox" data-light-source-toggle>
            <span>${localizeTracker('TimedEffects.Settings.LightSourceLabel')}</span>
          </label>
          <div class="expedition-timed-effect-light-fields" data-light-source-fields hidden>
            <label class="expedition-timed-effect-inline-field">
              <span>${localizeTracker('TimedEffects.Settings.BrightLightLabel')}</span>
              <input type="number" value="0" min="0" step="1" disabled>
            </label>
            <label class="expedition-timed-effect-inline-field">
              <span>${localizeTracker('TimedEffects.Settings.DimLightLabel')}</span>
              <input type="number" value="0" min="0" step="1" disabled>
            </label>
            <label class="expedition-timed-effect-inline-field">
              <span>${localizeTracker('TimedEffects.Settings.EmissionAngleLabel')}</span>
              <input type="number" value="360" min="1" max="360" step="1" disabled>
            </label>
          </div>
        </div>
      </td>
      <td class="template-actions-cell">
        <button type="button" data-action="remove-template">
          ${localizeTracker('TimedEffects.Settings.RemoveTemplate')}
        </button>
      </td>
    `;
    container.append(row);
    this.#reindexTemplateRows(container);
  }

  #reindexTemplateRows(container) {
    if (!container) return;
    const rows = container.querySelectorAll('[data-template-row]');
    rows.forEach((row, index) => {
      const idInput = row.querySelector('input[type="hidden"]');
      const nameInput = row.querySelector('input[data-template-field="name"]') ?? row.querySelector('td:nth-child(1) input[type="text"]');
      const iconInput = row.querySelector('input[data-template-field="iconPath"]') ?? row.querySelector('td:nth-child(2) input[type="text"]');
      const durationInput = row.querySelector('input[data-template-field="defaultDurationTurns"]') ?? row.querySelector('input[type="number"]');
      const isLightSourceInput = row.querySelector('input[data-light-source-toggle]');
      const brightInput = row.querySelector('input[data-template-field="brightLight"]');
      const dimInput = row.querySelector('input[data-template-field="dimLight"]');
      const angleInput = row.querySelector('input[data-template-field="emissionAngle"]');

      if (idInput) idInput.name = `templates.${index}.id`;
      if (nameInput) {
        nameInput.name = `templates.${index}.name`;
        nameInput.dataset.templateField = 'name';
      }
      if (iconInput) {
        iconInput.name = `templates.${index}.iconPath`;
        iconInput.dataset.templateField = 'iconPath';
      }
      if (durationInput) {
        durationInput.name = `templates.${index}.defaultDurationTurns`;
        durationInput.dataset.templateField = 'defaultDurationTurns';
      }
      if (isLightSourceInput) {
        isLightSourceInput.name = `templates.${index}.isLightSource`;
      }
      if (brightInput) brightInput.name = `templates.${index}.brightLight`;
      if (dimInput) dimInput.name = `templates.${index}.dimLight`;
      if (angleInput) angleInput.name = `templates.${index}.emissionAngle`;
    });
  }

  #syncAllLightSourceRows(rootElement) {
    rootElement?.querySelectorAll('[data-template-row]').forEach((row) => this.#syncLightSourceRow(row));
  }

  #syncLightSourceRow(row) {
    if (!row) return;
    const checkbox = row.querySelector('input[data-light-source-toggle]');
    const fields = row.querySelector('[data-light-source-fields]');
    const isEnabled = Boolean(checkbox?.checked);
    fields?.toggleAttribute('hidden', !isEnabled);
    fields?.querySelectorAll('input').forEach((input) => {
      input.disabled = !isEnabled;
    });
  }

  #syncWindowHeight(rootElement) {
    const position = this.position ?? {};
    const windowElement = rootElement?.closest('.window-app');
    const measuredHeight = windowElement?.scrollHeight ?? rootElement?.scrollHeight ?? position.height ?? 'auto';

    this.setPosition({
      height: Math.min(
        TIMED_EFFECT_TEMPLATE_CONFIG_MAX_HEIGHT,
        Math.max(320, Number(measuredHeight) || 320),
      ),
    });
  }
}

class TrackerPositionResetConfig extends FormApplication {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      id: `${MODULE_ID}-tracker-position-reset`,
      template: TRACKER_POSITION_RESET_TEMPLATE,
      title: localizeTracker('TrackerPosition.ResetMenuLabel'),
      width: 420,
      height: 'auto',
      closeOnSubmit: true,
      submitOnClose: false,
      classes: ['expedition-timed-effect-template-config'],
    });
  }

  async _updateObject() {
    await game.settings.set(MODULE_ID, TRACKER_POSITION_SETTING_KEY, {});
    ui.notifications.info(localizeTracker('TrackerPosition.ResetSuccess'));
    Hooks.callAll(EXPEDITION_TRACKER_UPDATE_HOOK);
  }
}

export function registerExpeditionTrackerSettings() {
  game.settings.register(MODULE_ID, TRACKER_ENABLED_SETTING_KEY, {
    name: 'EXPEDITION_TRACKER.Settings.Enable.Name',
    hint: 'EXPEDITION_TRACKER.Settings.Enable.Hint',
    scope: 'world',
    config: true,
    type: Boolean,
    default: false,
    onChange: () => Hooks.callAll(EXPEDITION_TRACKER_UPDATE_HOOK),
  });

  game.settings.register(MODULE_ID, TIMED_EFFECTS_ENABLED_SETTING_KEY, {
    name: 'EXPEDITION_TRACKER.Settings.TimedEffectsEnable.Name',
    hint: 'EXPEDITION_TRACKER.Settings.TimedEffectsEnable.Hint',
    scope: 'world',
    config: true,
    type: Boolean,
    default: false,
    onChange: () => Hooks.callAll(EXPEDITION_TRACKER_UPDATE_HOOK),
  });

  game.settings.register(MODULE_ID, TRACKER_SETTING_KEY, {
    scope: 'world',
    config: false,
    type: Object,
    default: getDefaultTrackerState(),
    onChange: () => Hooks.callAll(EXPEDITION_TRACKER_UPDATE_HOOK),
  });

  game.settings.register(MODULE_ID, TIMED_EFFECT_TEMPLATES_SETTING_KEY, {
    scope: 'world',
    config: false,
    type: Object,
    default: [],
    onChange: () => Hooks.callAll(EXPEDITION_TRACKER_UPDATE_HOOK),
  });

  game.settings.register(MODULE_ID, TRACKER_POSITION_SETTING_KEY, {
    scope: 'client',
    config: false,
    type: Object,
    default: {},
    onChange: () => Hooks.callAll(EXPEDITION_TRACKER_UPDATE_HOOK),
  });

  game.settings.registerMenu(MODULE_ID, TIMED_EFFECT_TEMPLATE_MENU_KEY, {
    name: 'EXPEDITION_TRACKER.TimedEffects.Settings.MenuLabel',
    label: 'EXPEDITION_TRACKER.TimedEffects.Settings.MenuAction',
    hint: 'EXPEDITION_TRACKER.TimedEffects.Settings.MenuHint',
    icon: 'fas fa-fire',
    type: TimedEffectTemplatesConfig,
    restricted: true,
  });

  game.settings.registerMenu(MODULE_ID, TRACKER_POSITION_RESET_MENU_KEY, {
    name: 'EXPEDITION_TRACKER.TrackerPosition.ResetMenuLabel',
    label: 'EXPEDITION_TRACKER.TrackerPosition.ResetMenuAction',
    hint: 'EXPEDITION_TRACKER.TrackerPosition.ResetMenuHint',
    icon: 'fas fa-up-down-left-right',
    type: TrackerPositionResetConfig,
    restricted: true,
  });
}
