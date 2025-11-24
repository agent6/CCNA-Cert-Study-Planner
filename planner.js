const STORAGE_KEY = 'ccnaPlannerState-v1';
const COURSE_PATH = 'course.md';
const DAY_MS = 86_400_000;
const DAY_LABELS_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const DAY_LABELS_LONG = [
  'Sunday',
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday',
];

const elements = {
  planForm: document.getElementById('planForm'),
  startInput: document.getElementById('startDate'),
  endInput: document.getElementById('endDate'),
  resetButton: document.getElementById('resetPlan'),
  editButton: document.getElementById('editPlan'),
  exportButton: document.getElementById('exportPlan'),
  plannerConfig: document.querySelector('.planner-config'),
  timelineRadios: () => elements.planForm?.elements?.timelineOption,
  studyDayInputs: () => document.querySelectorAll('input[name="studyDay"]'),
  studyDayFieldset: document.querySelector('.weekdays-fieldset'),
  studyDayError: document.getElementById('studyDaysError'),
  studyDaySummary: document.getElementById('studyDaysSummary'),
  dayGrid: document.querySelector('.day-grid'),
  cards: {
    yesterday: document.querySelector('.day-card[data-day="yesterday"]'),
    today: document.querySelector('.day-card[data-day="today"]'),
    tomorrow: document.querySelector('.day-card[data-day="tomorrow"]'),
  },
  stats: {
    progressValue: document.getElementById('progressValue'),
    progressDetail: document.getElementById('progressDetail'),
    progressBar: document.getElementById('progressBar'),
    tasksRemaining: document.getElementById('tasksRemaining'),
    tasksPerDayHint: document.getElementById('tasksPerDayHint'),
    daysRemaining: document.getElementById('daysRemaining'),
    deadlineLabel: document.getElementById('deadlineLabel'),
  },
  month: {
    label: document.getElementById('monthLabel'),
    grid: document.getElementById('monthGrid'),
    navButtons: document.querySelectorAll('[data-month-step]'),
  },
  template: document.getElementById('taskTemplate'),
};

let courseTasks = [];
let state = null;
let isEditingPlan = false;
let dragState = null;

init();

async function init() {
  courseTasks = await loadCourse();
  state = hydrateState(courseTasks);

  if (!state.viewMonth) {
    state.viewMonth = toIso(startOfMonth(startOfToday()));
  } else {
    const monthDate = fromIso(state.viewMonth) || startOfToday();
    state.viewMonth = toIso(startOfMonth(monthDate));
  }

  if (state.planCommitted && state.startDate) {
    const startDate = fromIso(state.startDate);
    const deadline = resolveDeadline(state);
    const hasSchedule = state.tasks.some((task) => task.scheduledDate);
    if (!hasSchedule && startDate && deadline) {
      rebuildScheduleAssignments(startDate, deadline);
      saveState();
    }
  }

  isEditingPlan = !state.planCommitted;
  bindEvents();
  render();
}

function bindEvents() {
  elements.planForm?.addEventListener('submit', handlePlanSubmit);
  elements.resetButton?.addEventListener('click', handleResetPlan);
  elements.dayGrid?.addEventListener('change', handleTaskToggle);
  elements.dayGrid?.addEventListener('dragstart', handleDragStart);
  elements.dayGrid?.addEventListener('dragend', handleDragEnd);
  elements.dayGrid?.addEventListener('dragover', handleDragOver);
  elements.dayGrid?.addEventListener('drop', handleDrop);
  elements.dayGrid?.addEventListener('dragleave', handleDragLeave);
  elements.dayGrid?.addEventListener('dragenter', handleDragEnter);

  elements.startInput?.addEventListener('change', (e) => {
    syncDeadlineInput();
    e.target.blur();
    document.activeElement.blur();
  });

  elements.endInput?.addEventListener('change', (e) => {
    const radios = elements.timelineRadios();
    if (radios) {
      radios.value = 'custom';
    }
    e.target.blur();
    document.activeElement.blur();
  });

  elements.endInput?.addEventListener('input', () => {
    const radios = elements.timelineRadios();
    if (radios) {
      radios.value = 'custom';
    }
  });

  const radios = toRadioArray(elements.timelineRadios());
  radios.forEach((radio) =>
    radio.addEventListener('change', () => {
      syncDeadlineInput();
    }),
  );

  elements.editButton?.addEventListener('click', handleEditPlan);
  elements.exportButton?.addEventListener('click', handleExportPlan);

  const studyDayInputs = Array.from(elements.studyDayInputs() || []);
  studyDayInputs.forEach((input) =>
    input.addEventListener('change', handleStudyDayChange),
  );

  elements.month.navButtons?.forEach((button) =>
    button.addEventListener('click', handleMonthNavigation),
  );
}

async function loadCourse() {
  try {
    const response = await fetch(COURSE_PATH, { cache: 'no-cache' });
    if (!response.ok) {
      throw new Error(`Course download failed with status ${response.status}`);
    }
    const text = await response.text();
    return parseCourse(text);
  } catch (error) {
    console.error('Failed to load course outline', error);
    displayCourseLoadError(error);
    return [];
  }
}

function parseCourse(raw) {
  const lines = raw
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);

  const tasks = [];
  let currentSkill = '';
  const idTracker = new Set();

  for (const line of lines) {
    if (/^skill\s+\d+/i.test(line)) {
      currentSkill = line;
      continue;
    }
    if (!/^(lesson|lab|quiz)/i.test(line)) {
      continue;
    }

    const typeMatch = line.match(/^(Lesson|Lab|Quiz)/i);
    const type = typeMatch ? capitalise(typeMatch[0]) : 'Task';

    const id = createTaskId(currentSkill, line, idTracker);

    tasks.push({
      id,
      title: line,
      type,
      skill: currentSkill,
      completed: false,
      completedOn: null,
      scheduledDate: null,
      order: tasks.length,
    });
  }

  return tasks;
}

function createTaskId(skill, line, registry) {
  const base = `${skill || 'general'}-${line}`
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 64);
  let id = base || `task-${registry.size + 1}`;
  let counter = 1;
  while (registry.has(id)) {
    id = `${base}-${counter++}`;
  }
  registry.add(id);
  return id;
}

function hydrateState(parsedTasks) {
  const defaultState = {
    startDate: '',
    timelineWeeks: null,
    timelineMode: '',
    customEndDate: '',
    planCommitted: false,
    viewMonth: '',
    studyDays: defaultStudyDays(),
    tasks: parsedTasks.map((task) => ({ ...task })),
  };

  const stored = localStorage.getItem(STORAGE_KEY);
  if (!stored) {
    return defaultState;
  }

  try {
    const saved = JSON.parse(stored);
    const savedTaskMap = new Map(
      (saved.tasks || []).map((task) => [task.id, task]),
    );

    const mergedTasks = parsedTasks.map((task) => {
      const savedTask = savedTaskMap.get(task.id);
      if (!savedTask) {
        return { ...task };
      }
      return {
        ...task,
        completed: Boolean(savedTask.completed),
        completedOn: savedTask.completedOn || null,
        scheduledDate: savedTask.scheduledDate || null,
      };
    });

    const viewMonthDate = saved.viewMonth
      ? fromIso(saved.viewMonth)
      : null;

    return {
      ...defaultState,
      ...saved,
      viewMonth: viewMonthDate
        ? toIso(startOfMonth(viewMonthDate))
        : defaultState.viewMonth,
      studyDays: sanitiseStudyDays(saved.studyDays),
      tasks: mergedTasks,
    };
  } catch (error) {
    console.warn('Failed to hydrate saved state. Falling back to defaults.', error);
    return defaultState;
  }
}

function saveState() {
  if (!state) return;
  const serialisable = {
    startDate: state.startDate,
    timelineWeeks: state.timelineWeeks,
    timelineMode: state.timelineMode,
    customEndDate: state.customEndDate,
    planCommitted: state.planCommitted,
    viewMonth: state.viewMonth,
    studyDays: state.studyDays,
    tasks: state.tasks.map(
      ({
        id,
        title,
        skill,
        type,
        order,
        completed,
        completedOn,
        scheduledDate,
      }) => ({
        id,
        title,
        skill,
        type,
        order,
        completed,
        completedOn,
        scheduledDate,
      }),
    ),
  };
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(serialisable));
  } catch (error) {
    console.warn('Unable to persist planner state.', error);
  }
}

function handlePlanSubmit(event) {
  event.preventDefault();
  const startValue = elements.startInput?.value;
  const radioList = elements.timelineRadios();
  const selectedTimeline = radioList ? radioList.value : '';
  const radios = toRadioArray(radioList);
  const endValue = elements.endInput?.value || '';
  const selectedStudyDays = getSelectedStudyDays();

  if (!startValue) {
    elements.startInput?.reportValidity();
    return;
  }

  if (!selectedTimeline) {
    radios.forEach((radio) => radio.setCustomValidity('Please choose a timeline.'));
    radios[0]?.reportValidity();
    radios.forEach((radio) => radio.setCustomValidity(''));
    return;
  }

  if (!selectedStudyDays.length) {
    showStudyDayError('Select at least one study day.');
    const studyInputs = elements.studyDayInputs();
    if (studyInputs && studyInputs.length) {
      studyInputs[0].focus();
    }
    return;
  }

  let timelineWeeks = null;
  let customEndDate = '';

  if (selectedTimeline === 'custom') {
    if (!endValue) {
      elements.endInput?.setCustomValidity('Pick a deadline.');
      elements.endInput?.reportValidity();
      elements.endInput?.setCustomValidity('');
      return;
    }
    if (compareIsoDates(endValue, startValue) < 0) {
      elements.endInput?.setCustomValidity('Deadline must be after the start date.');
      elements.endInput?.reportValidity();
      elements.endInput?.setCustomValidity('');
      return;
    }
    customEndDate = endValue;
  } else {
    timelineWeeks = Number(selectedTimeline);
    customEndDate = '';
  }

  state.startDate = startValue;
  state.timelineMode = selectedTimeline;
  state.timelineWeeks = timelineWeeks;
  state.customEndDate = customEndDate;
  state.planCommitted = true;
  state.studyDays = selectedStudyDays;
  state.viewMonth = toIso(startOfMonth(fromIso(startValue) || startOfToday()));

  const startDate = fromIso(startValue);
  const deadline = resolveDeadline(state);
  rebuildScheduleAssignments(startDate, deadline);

  isEditingPlan = false;
  clearStudyDayError();
  updateStudyDaySummary(selectedStudyDays);
  syncDeadlineInput();
  saveState();
  render();
}

function handleStudyDayChange() {
  const selected = getSelectedStudyDays();
  if (selected.length) {
    clearStudyDayError();
  } else {
    showStudyDayError('Select at least one study day.');
  }
  updateStudyDaySummary(selected);
}

function handleEditPlan() {
  if (!state?.planCommitted) {
    return;
  }
  isEditingPlan = true;
  renderPlanVisibility(true);
  clearStudyDayError();
  updateStudyDaySummary(state.studyDays);
  requestAnimationFrame(() => {
    elements.startInput?.focus();
  });
}

function handleResetPlan(e) {
  if (e) e.preventDefault();
  if (!state) return;
  const hasProgress = state.tasks.some((task) => task.completed);
  if (!hasProgress) {
    return;
  }
  const shouldReset = window.confirm(
    'Reset all completed tasks? This does not change your dates.',
  );
  if (!shouldReset) {
    return;
  }
  state.tasks.forEach((task) => {
    task.completed = false;
    task.completedOn = null;
  });
  saveState();
  render();
}

async function handleExportPlan() {
  if (!state?.planCommitted) {
    window.alert('Save your plan before exporting.');
    return;
  }

  const assets = {};

  try {
    const logo = await loadPdfLogo('assets/academy.png');
    if (logo) {
      assets[logo.name] = logo;
    }
  } catch (error) {
    console.error('Failed to load PDF logo', error);
  }

  const blocks = buildPlanBlocksForPdf(state, assets);
  try {
    const pdfBytes = createPlanPdf(blocks, assets);
    const fileName = createPdfFileName(state);
    downloadBlob(pdfBytes, fileName);
  } catch (error) {
    console.error('Failed to export PDF', error);
    window.alert('PDF export failed. Please try again.');
  }
}

function buildPlanBlocksForPdf(localState, assets = {}) {
  const theme = getPdfTheme();
  const blocks = [];
  const addLine = (text, config = {}) =>
    blocks.push({
      type: 'line',
      text,
      font: 'regular',
      fontSize: 12,
      color: theme.text,
      indent: 0,
      lineHeight: 1.4,
      before: 0,
      after: 0,
      align: 'left',
      ...config,
    });
  const addRule = (config = {}) =>
    blocks.push({
      type: 'rule',
      thickness: 1,
      color: theme.border,
      before: 10,
      after: 10,
      width: null,
      align: 'left',
      ...config,
    });
  const addTask = (text, config = {}) =>
    blocks.push({
      type: 'task',
      text,
      checked: Boolean(config.checked),
      indent: config.indent ?? 18,
      fontSize: config.fontSize ?? 11,
      lineHeight: config.lineHeight ?? 1.35,
      before: config.before ?? 0,
      after: config.after ?? 3,
      color: config.color ?? theme.text,
      borderColor: config.borderColor ?? theme.border,
      checkColor: config.checkColor ?? theme.accent,
    });

  blocks.push({
    type: 'masthead',
    logoKey: assets.logo ? assets.logo.name : null,
    logoWidth: assets.logo ? assets.logo.displayWidth : 0,
    logoHeight: assets.logo ? assets.logo.displayHeight : 0,
    gap: assets.logo ? 18 : 0,
    title: 'Academy CCNA Study Plan',
    titleFontSize: 22,
    titleColor: theme.text,
    subtitle: 'Powered by NetworkChuck',
    subtitleFontSize: 12,
    subtitleColor: theme.accent,
    after: 18,
  });

  addRule({ color: theme.accent, thickness: 2, width: 360, align: 'left', after: 18 });

  addLine('Learner Summary', {
    font: 'bold',
    fontSize: 12,
    color: theme.muted,
    after: 10,
  });

  const startDate = localState.startDate ? fromIso(localState.startDate) : null;
  const deadlineDate = resolveDeadline(localState);
  const startLabel = startDate ? formatLongDate(startDate) : 'Not set';
  const deadlineLabel = deadlineDate ? formatLongDate(deadlineDate) : 'Not set';

  const summaryItems = [
    { label: 'Start date', value: startLabel },
    { label: 'Target deadline', value: deadlineLabel },
  ];

  summaryItems.forEach(({ label, value }) => {
    addLine(label.toUpperCase(), {
      font: 'bold',
      fontSize: 10,
      color: theme.accentMuted,
    });
    addLine(value, {
      indent: 18,
      fontSize: 12,
      color: theme.text,
      after: 10,
    });
  });

  addRule({ color: theme.border, thickness: 1, after: 12 });

  const schedule = buildScheduleMap();
  const dayKeys = Array.from(schedule.keys()).sort(compareIsoDates);
  const unscheduled = localState.tasks
    .filter((task) => !task.scheduledDate)
    .sort((a, b) => a.order - b.order);

  addLine('Scheduled Tasks', {
    font: 'bold',
    fontSize: 14,
    color: theme.text,
    after: 8,
  });

  if (dayKeys.length) {
    dayKeys.forEach((dateKey) => {
      const date = fromIso(dateKey);
      if (!date) {
        return;
      }
      const tasksForDay = schedule.get(dateKey) || [];
      const doneCount = tasksForDay.filter((task) => task.completed).length;
      const dueCount = tasksForDay.length - doneCount;
      const summaryParts = [];
      summaryParts.push(
        `${tasksForDay.length} task${tasksForDay.length === 1 ? '' : 's'}`,
      );
      if (dueCount > 0) {
        summaryParts.push(`${dueCount} due`);
      }
      if (doneCount > 0) {
        summaryParts.push(`${doneCount} done`);
      }

      addLine(formatLongDate(date), {
        font: 'bold',
        fontSize: 12,
        color: theme.accent,
        before: 6,
      });

      addLine(summaryParts.join(' | '), {
        indent: 18,
        fontSize: 11,
        color: theme.muted,
        after: tasksForDay.length ? 4 : 8,
      });

      if (!tasksForDay.length) {
        addLine('No tasks scheduled for this day.', {
          indent: 26,
          fontSize: 11,
          color: theme.muted,
          after: 8,
        });
        return;
      }

      tasksForDay.forEach((task, index) => {
        addTask(task.title, {
          indent: 26,
          checked: task.completed,
          after: index === tasksForDay.length - 1 ? 10 : 3,
        });
      });
    });
  } else {
    addLine('No tasks have been scheduled yet.', {
      indent: 18,
      fontSize: 11,
      color: theme.muted,
    });
  }

  if (unscheduled.length) {
    addRule({ color: theme.border, thickness: 1, after: 12 });
    addLine('Unscheduled Tasks', {
      font: 'bold',
      fontSize: 13,
      color: theme.text,
      after: 8,
    });
    unscheduled.forEach((task, index) => {
      addTask(task.title, {
        indent: 24,
        checked: false,
        after: index === unscheduled.length - 1 ? 8 : 3,
      });
    });
  }

  addRule({ color: theme.border, thickness: 1, after: 12 });
  addLine('Stay focused - review this planner each study day and check off completed work.', {
    fontSize: 10,
    color: theme.muted,
  });

  return blocks;
}

function createPlanPdf(blocks, assets = {}) {
  const options = {
    pageWidth: 612,
    pageHeight: 792,
    margin: 54,
  };
  const pages = layoutPdfBlocks(blocks, options, assets);
  return buildPdfDocument(pages, options, assets);
}

function createPdfFileName(localState) {
  const base = localState?.startDate || toIso(startOfToday()) || 'plan';
  return `academy-ccna-study-plan-${base}.pdf`;
}

function downloadBlob(bytes, fileName) {
  const blob = new Blob([bytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function layoutPdfBlocks(blocks, options, assets = {}) {
  const contentWidth = options.pageWidth - options.margin * 2;
  const contentHeight = options.pageHeight - options.margin * 2;
  const topBaseline = options.pageHeight - options.margin;
  const pages = [];
  let cursor = 0;
  let currentPage = { ops: [] };

  const newPage = () => {
    if (currentPage.ops.length) {
      pages.push(currentPage);
    }
    currentPage = { ops: [] };
    cursor = 0;
  };

  const ensureSpace = (amount) => {
    if (cursor + amount > contentHeight && currentPage.ops.length) {
      newPage();
    }
  };

  blocks.forEach((block) => {
    if (block.type === 'masthead') {
      const before = block.before ?? 0;
      const after = block.after ?? 18;
      const gap = block.gap ?? 18;
      const logoAsset = block.logoKey ? assets[block.logoKey] : null;
      let logoWidth = 0;
      let logoHeight = 0;
      if (logoAsset) {
        const requestedWidth = block.logoWidth || logoAsset.displayWidth || 0;
        logoWidth = Math.min(requestedWidth, contentWidth * 0.35);
        const baseWidth = logoAsset.displayWidth || (logoWidth || 1);
        const ratio = baseWidth ? logoWidth / baseWidth : 1;
        logoHeight = (logoAsset.displayHeight || logoWidth) * ratio;
      }
      const titleFontSize = block.titleFontSize ?? 22;
      const subtitleFontSize = block.subtitleFontSize ?? 12;
      const lineGap = block.lineGap ?? 6;
      const textHeight = titleFontSize + subtitleFontSize + lineGap;
      const blockHeight = Math.max(logoHeight, textHeight);
      ensureSpace(before + blockHeight + after);
      cursor += before;
      const baseY = topBaseline - cursor;

      const logoOffset = Math.max((blockHeight - logoHeight) / 2, 0);

      if (logoAsset && logoWidth > 0) {
        currentPage.ops.push({
          type: 'image',
          assetKey: block.logoKey,
          x: options.margin,
          y: baseY - logoOffset - logoHeight,
          width: logoWidth,
          height: logoHeight,
        });
      }

      const textX = options.margin + (logoWidth > 0 ? logoWidth + gap : 0);
      const textOffset = Math.max((blockHeight - textHeight) / 2, 0);
      let textY = baseY - textOffset - titleFontSize;
      currentPage.ops.push({
        type: 'text',
        text: block.title,
        font: 'bold',
        fontSize: titleFontSize,
        color: hexToRgbArray(block.titleColor || '#000000'),
        x: textX,
        y: textY,
      });
      textY -= subtitleFontSize + lineGap;
      currentPage.ops.push({
        type: 'text',
        text: block.subtitle,
        font: 'bold',
        fontSize: subtitleFontSize,
        color: hexToRgbArray(block.subtitleColor || '#000000'),
        x: textX,
        y: textY,
      });

      cursor += blockHeight + after;
      return;
    }

    if (block.type === 'image') {
      const asset = assets[block.assetKey];
      if (!asset) {
        return;
      }
      const before = block.before ?? 0;
      const after = block.after ?? 12;
      const indent = block.indent ?? 0;
      const availableWidth = Math.max(contentWidth - indent, 24);
      const width = Math.min(block.width ?? asset.displayWidth, availableWidth);
      const aspectRatio = asset.displayHeight / asset.displayWidth || 1;
      const height = block.height
        ? Math.min(block.height, contentHeight)
        : width * aspectRatio;
      ensureSpace(before + height + after);
      cursor += before;
      const x =
        block.align === 'center'
          ? options.margin + (contentWidth - width) / 2
          : options.margin + indent;
      const y = topBaseline - cursor - height;
      currentPage.ops.push({
        type: 'image',
        assetKey: block.assetKey,
        x,
        y,
        width,
        height,
      });
      cursor += height + after;
      return;
    }

    if (block.type === 'rule') {
      const before = block.before ?? 10;
      const after = block.after ?? 10;
      const thickness = block.thickness ?? 1;
      const width =
        block.width && block.width > 0 ? Math.min(block.width, contentWidth) : contentWidth;
      ensureSpace(before + thickness + after);
      cursor += before;
      const xOffset =
        block.align === 'center' ? options.margin + (contentWidth - width) / 2 : options.margin;
      const y = topBaseline - cursor - thickness / 2;
      currentPage.ops.push({
        type: 'rule',
        x: xOffset,
        y,
        width,
        thickness,
        color: hexToRgbArray(block.color),
      });
      cursor += thickness + after;
      return;
    }

    if (block.type === 'task') {
      const before = block.before ?? 0;
      const after = block.after ?? 3;
      const indent = block.indent ?? 18;
      const fontSize = block.fontSize ?? 11;
      const lineHeight = (block.lineHeight ?? 1.35) * fontSize;
      const boxSize = block.boxSize ?? 10;
      const gap = block.gap ?? 6;
      const maxWidth = contentWidth - indent - boxSize - gap;
      const wrapped = wrapLineForPdf(block.text, {
        maxWidth,
        fontSize,
      });

      ensureSpace(before);
      cursor += before;

      wrapped.forEach((part, lineIndex) => {
        ensureSpace(lineHeight);
        const lineMiddle = topBaseline - cursor - lineHeight / 2;
        const baseline = lineMiddle + fontSize * 0.35;
        const textX = options.margin + indent + boxSize + gap;

        if (lineIndex === 0) {
          const boxY = baseline - fontSize*.1;
          currentPage.ops.push({
            type: 'checkbox',
            x: options.margin + indent,
            y: boxY,
            size: boxSize,
            borderColor: hexToRgbArray(block.borderColor ?? '#c59f7a'),
            checked: block.checked,
            checkColor: hexToRgbArray(block.checkColor ?? '#d86a28'),
          });
        }

        currentPage.ops.push({
          type: 'text',
          text: part,
          font: block.font === 'bold' ? 'bold' : 'regular',
          fontSize,
          color: hexToRgbArray(block.color ?? '#000000'),
          x: textX,
          y: baseline,
        });
        cursor += lineHeight;
      });

      cursor += after;
      return;
    }

    if (block.type !== 'line') {
      return;
    }

    const before = block.before ?? 0;
    const after = block.after ?? 0;
    const indent = block.indent ?? 0;
    const fontSize = block.fontSize ?? 12;
    const lineHeight = (block.lineHeight ?? 1.4) * fontSize;
    const maxWidth =
      block.maxWidth && block.maxWidth > 0
        ? Math.min(block.maxWidth, contentWidth - indent)
        : contentWidth - indent;
    const align = block.align || 'left';
    const approxCharWidth = fontSize * 0.53;

    const wrapped = wrapLineForPdf(block.text, {
      maxWidth,
      fontSize,
    });

    ensureSpace(before);
    cursor += before;

    wrapped.forEach((part) => {
      ensureSpace(lineHeight);
      const y = topBaseline - cursor;
      let textX = options.margin + indent;
      if (align === 'center') {
        const estimatedWidth = Math.min(part.length * approxCharWidth, maxWidth);
        textX = options.margin + indent + Math.max((maxWidth - estimatedWidth) / 2, 0);
      } else if (align === 'right') {
        const estimatedWidth = Math.min(part.length * approxCharWidth, maxWidth);
        textX = options.margin + indent + Math.max(maxWidth - estimatedWidth, 0);
      }
      currentPage.ops.push({
        type: 'text',
        text: part,
        font: block.font === 'bold' ? 'bold' : 'regular',
        fontSize,
        color: hexToRgbArray(block.color),
        x: textX,
        y,
      });
      cursor += lineHeight;
    });

    cursor += after;
  });

  if (currentPage.ops.length || pages.length === 0) {
    pages.push(currentPage);
  }

  return pages;
}

function wrapLineForPdf(text, { maxWidth, fontSize }) {
  if (!text) return [''];
  const approximateCharWidth = fontSize * 0.53;
  const maxChars = Math.max(Math.floor(maxWidth / approximateCharWidth), 8);
  const words = text.split(/\s+/);
  const lines = [];
  let buffer = '';

  const flush = () => {
    if (buffer) {
      lines.push(buffer);
      buffer = '';
    }
  };

  words.forEach((word) => {
    if (!word) return;
    const candidate = buffer ? `${buffer} ${word}` : word;
    if (candidate.length <= maxChars) {
      buffer = candidate;
      return;
    }
    flush();
    if (word.length > maxChars) {
      let remaining = word;
      while (remaining.length > maxChars) {
        lines.push(remaining.slice(0, maxChars));
        remaining = remaining.slice(maxChars);
      }
      buffer = remaining;
    } else {
      buffer = word;
    }
  });

  flush();

  if (!lines.length) {
    lines.push('');
  }

  return lines;
}

function buildPdfDocument(pages, options, assets = {}) {
  const encoder = new TextEncoder();
  const objects = [];

  const reserveObject = () => {
    objects.push(null);
    return objects.length;
  };

  const setObject = (id, content) => {
    objects[id - 1] = content;
  };

  const catalogId = reserveObject();
  const pagesId = reserveObject();
  const fontRegularId = reserveObject();
  const fontBoldId = reserveObject();

  const imageEntries = [];
  Object.values(assets).forEach((asset) => {
    if (asset?.colorData?.length) {
      const objectId = reserveObject();
      const maskId = reserveObject();
      const resourceName = `Im${imageEntries.length + 1}`;
      imageEntries.push({ asset, objectId, maskId, resourceName });
    } else if (asset?.data?.length) {
      const objectId = reserveObject();
      const resourceName = `Im${imageEntries.length + 1}`;
      imageEntries.push({ asset, objectId, resourceName, maskId: null });
    }
  });

  const pageEntries = pages.map(() => ({
    contentId: reserveObject(),
    pageId: reserveObject(),
  }));

  setObject(fontRegularId, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
  setObject(fontBoldId, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>');

  imageEntries.forEach(({ asset, objectId, maskId }) => {
    if (asset.colorData?.length && asset.alphaData?.length) {
      const maskStream = buildMaskImageObject(asset);
      setObject(maskId, maskStream);
      const stream = buildImageXObject(asset, maskId);
      setObject(objectId, stream);
    } else {
      const stream = buildImageXObject(asset);
      setObject(objectId, stream);
    }
  });

  const imageResources = new Map();
  imageEntries.forEach(({ asset, resourceName, objectId }) => {
    imageResources.set(asset.name, { resourceName, objectId });
  });

  pageEntries.forEach((entry, index) => {
    const stream = buildPageContentStream(pages[index].ops, imageResources);
    const length = encoder.encode(stream).length;
    const content = `<< /Length ${length} >>\nstream\n${stream}\nendstream`;
    setObject(entry.contentId, content);

    const xObjectSection =
      imageEntries.length > 0
        ? ` /XObject << ${imageEntries
            .map(({ resourceName, objectId }) => `/${resourceName} ${objectId} 0 R`)
            .join(' ')} >>`
        : '';

    const pageContent = [
      '<<',
      ' /Type /Page',
      ` /Parent ${pagesId} 0 R`,
      ` /MediaBox [0 0 ${options.pageWidth} ${options.pageHeight}]`,
      ` /Resources << /Font << /F1 ${fontRegularId} 0 R /F2 ${fontBoldId} 0 R >>${xObjectSection} >>`,
      ` /Contents ${entry.contentId} 0 R`,
      '>>',
    ].join('\n');
    setObject(entry.pageId, pageContent);
  });

  const kids = pageEntries.map((entry) => `${entry.pageId} 0 R`).join(' ');
  const pagesContent = ['<<', ' /Type /Pages', ` /Count ${pages.length}`, ` /Kids [ ${kids} ]`, '>>'].join(
    '\n',
  );
  setObject(pagesId, pagesContent);

  const catalogContent = ['<<', ' /Type /Catalog', ` /Pages ${pagesId} 0 R`, '>>'].join('\n');
  setObject(catalogId, catalogContent);

  const header = '%PDF-1.4\n';
  const xrefEntries = ['0000000000 65535 f \n'];
  let body = '';
  let offset = header.length;

  objects.forEach((content, index) => {
    if (content == null) {
      throw new Error('Incomplete PDF object graph');
    }
    const objectId = index + 1;
    const chunk = `${objectId} 0 obj\n${content}\nendobj\n`;
    xrefEntries.push(`${String(offset).padStart(10, '0')} 00000 n \n`);
    body += chunk;
    offset += chunk.length;
  });

  const startXref = offset;
  const xref = `xref\n0 ${objects.length + 1}\n${xrefEntries.join('')}`;
  const trailer = `trailer\n<< /Size ${objects.length + 1} /Root ${catalogId} 0 R >>\nstartxref\n${startXref}\n%%EOF`;
  const pdfString = header + body + xref + trailer;
  return encoder.encode(pdfString);
}

function buildPageContentStream(ops, imageResources) {
  if (!ops.length) {
    return ' ';
  }
  const commands = [];
  ops.forEach((op) => {
    if (op.type === 'text') {
      commands.push('BT');
      commands.push(`/F${op.font === 'bold' ? '2' : '1'} ${op.fontSize} Tf`);
      commands.push(`${op.color.join(' ')} rg`);
      commands.push(`1 0 0 1 ${op.x.toFixed(2)} ${op.y.toFixed(2)} Tm`);
      commands.push(`(${escapePdfText(op.text)}) Tj`);
      commands.push('ET');
    } else if (op.type === 'rule') {
      commands.push(`${op.color.join(' ')} RG`);
      commands.push(`${op.color.join(' ')} rg`);
      commands.push(`${op.thickness.toFixed(2)} w`);
      commands.push(
        `${op.x.toFixed(2)} ${op.y.toFixed(2)} m ${(op.x + op.width).toFixed(2)} ${op.y.toFixed(2)} l S`,
      );
    } else if (op.type === 'checkbox') {
      commands.push('q');
      commands.push(`${op.borderColor.join(' ')} RG`);
      commands.push('1 1 1 rg');
      commands.push(`${op.x.toFixed(2)} ${op.y.toFixed(2)} ${op.size.toFixed(2)} ${op.size.toFixed(2)} re B`);
      if (op.checked) {
        const startX = op.x + op.size * 0.2;
        const startY = op.y + op.size * 0.55;
        const midX = op.x + op.size * 0.45;
        const midY = op.y + op.size * 0.2;
        const endX = op.x + op.size * 0.85;
        const endY = op.y + op.size * 0.8;
        commands.push(`${op.checkColor.join(' ')} RG`);
        commands.push(`${op.checkColor.join(' ')} rg`);
        commands.push('1.4 w');
        commands.push(
          `${startX.toFixed(2)} ${startY.toFixed(2)} m ${midX.toFixed(2)} ${midY.toFixed(2)} l ${endX.toFixed(
            2,
          )} ${endY.toFixed(2)} l S`,
        );
      }
      commands.push('Q');
    } else if (op.type === 'image') {
      const resource = imageResources?.get(op.assetKey);
      if (!resource) {
        return;
      }
      commands.push('q');
      commands.push(
        `${op.width.toFixed(2)} 0 0 ${op.height.toFixed(2)} ${op.x.toFixed(2)} ${op.y.toFixed(2)} cm`,
      );
      commands.push(`/${resource.resourceName} Do`);
      commands.push('Q');
    }
  });
  return commands.join('\n');
}

function getPdfTheme() {
  return {
    text: '#121317',
    muted: '#525663',
    accent: '#d86a28',
    accentMuted: '#b65b24',
    accentSecondary: '#7a9c51',
    border: '#e4d9cb',
  };
}

function buildImageXObject(asset, maskRefId) {
  if (asset.colorData?.length && asset.alphaData?.length) {
    const encoded = ascii85Encode(asset.colorData);
    const header = [
      '<<',
      ' /Type /XObject',
      ' /Subtype /Image',
      ` /Width ${asset.pixelWidth}`,
      ` /Height ${asset.pixelHeight}`,
      ' /ColorSpace /DeviceRGB',
      ' /BitsPerComponent 8',
      ' /Filter /ASCII85Decode',
      ` /Length ${encoded.length}`,
      maskRefId ? ` /SMask ${maskRefId} 0 R` : '',
      '>>',
    ]
      .filter(Boolean)
      .join('\n');
    return `${header}\nstream\n${encoded}\nendstream`;
  }

  const encoded = ascii85Encode(asset.data);
  const header = [
    '<<',
    ' /Type /XObject',
    ' /Subtype /Image',
    ` /Width ${asset.pixelWidth}`,
    ` /Height ${asset.pixelHeight}`,
    ' /ColorSpace /DeviceRGB',
    ' /BitsPerComponent 8',
    ' /Filter [/ASCII85Decode /DCTDecode]',
    ` /Length ${encoded.length}`,
    '>>',
  ].join('\n');
  return `${header}\nstream\n${encoded}\nendstream`;
}

function buildMaskImageObject(asset) {
  const encoded = ascii85Encode(asset.alphaData);
  const header = [
    '<<',
    ' /Type /XObject',
    ' /Subtype /Image',
    ` /Width ${asset.pixelWidth}`,
    ` /Height ${asset.pixelHeight}`,
    ' /ColorSpace /DeviceGray',
    ' /BitsPerComponent 8',
    ' /Filter /ASCII85Decode',
    ` /Length ${encoded.length}`,
    '>>',
  ].join('\n');
  return `${header}\nstream\n${encoded}\nendstream`;
}

function hexToRgbArray(hex = '#000000') {
  const normalised = hex.replace('#', '');
  const chunk = normalised.length === 3 ? normalised.split('').map((ch) => ch + ch) : normalised.match(/.{1,2}/g);
  if (!chunk) {
    return [0, 0, 0];
  }
  return chunk
    .slice(0, 3)
    .map((value) => Math.min(parseInt(value, 16), 255) / 255)
    .map((value) => Number(value.toFixed(4)));
}

function escapePdfText(text) {
  return text
    .replace(/\\/g, '\\\\')
    .replace(/\(/g, '\\(')
    .replace(/\)/g, '\\)')
    .replace(/\r/g, '')
    .replace(/\n/g, ' ');
}

async function loadPdfLogo(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.crossOrigin = 'anonymous';
    image.onload = () => {
      const maxWidth = 220;
      const scale = Math.min(1, maxWidth / image.naturalWidth || 1);
      const targetWidth = Math.max(1, Math.round(image.naturalWidth * scale));
      const targetHeight = Math.max(1, Math.round(image.naturalHeight * scale));

      const canvas = document.createElement('canvas');
      canvas.width = targetWidth;
      canvas.height = targetHeight;
      const context = canvas.getContext('2d');
      if (!context) {
        reject(new Error('Canvas context unavailable.'));
        return;
      }
      context.drawImage(image, 0, 0, targetWidth, targetHeight);
    try {
      const imageData = context.getImageData(0, 0, targetWidth, targetHeight);
      const rgba = imageData.data;
      const pixelCount = targetWidth * targetHeight;
      const colorData = new Uint8Array(pixelCount * 3);
      const alphaData = new Uint8Array(pixelCount);
      for (let i = 0, c = 0, a = 0; i < rgba.length; i += 4) {
        colorData[c++] = rgba[i];
        colorData[c++] = rgba[i + 1];
        colorData[c++] = rgba[i + 2];
        alphaData[a++] = rgba[i + 3];
      }
      resolve({
        name: 'logo',
        colorData,
        alphaData,
        pixelWidth: targetWidth,
        pixelHeight: targetHeight,
        displayWidth: pxToPoints(targetWidth),
        displayHeight: pxToPoints(targetHeight),
      });
    } catch (canvasError) {
      reject(canvasError);
    }
    };
    image.onerror = () => reject(new Error('Failed to load logo image.'));
    image.src = src;
  });
}

function ascii85Encode(bytes) {
  let output = '';
  let i = 0;
  const length = bytes.length;
  while (i < length) {
    const remaining = length - i;
    const b0 = bytes[i];
    const b1 = remaining > 1 ? bytes[i + 1] : 0;
    const b2 = remaining > 2 ? bytes[i + 2] : 0;
    const b3 = remaining > 3 ? bytes[i + 3] : 0;
    const value = b0 * 16777216 + b1 * 65536 + b2 * 256 + b3;

    if (remaining >= 4 && value === 0) {
      output += 'z';
      i += 4;
      continue;
    }

    const chars = new Array(5);
    let temp = value;
    for (let j = 4; j >= 0; j -= 1) {
      chars[j] = (temp % 85) + 33;
      temp = Math.floor(temp / 85);
    }

    let chunk = String.fromCharCode(chars[0], chars[1], chars[2], chars[3], chars[4]);
    if (remaining < 4) {
      chunk = chunk.slice(0, remaining + 1);
    }

    output += chunk;
    i += 4;
  }

  output += '~>';
  return output;
}

function pxToPoints(px) {
  return Math.round(((px * 72) / 96) * 100) / 100;
}

function handleTaskToggle(event) {
  if (event.target.type !== 'checkbox') return;
  const taskItem = event.target.closest('.task-item');
  if (!taskItem) return;
  const card = event.target.closest('.day-card');
  if (!card) return;

  const taskId = taskItem.dataset.taskId;
  const dateKey = card.dataset.dateKey;
  const task = state.tasks.find((item) => item.id === taskId);
  if (!task) return;

  const checked = event.target.checked;
  task.completed = checked;
  task.completedOn = checked ? dateKey : null;

  saveState();
  render();
}

function handleMonthNavigation(event) {
  const { monthStep } = event.currentTarget.dataset;
  if (!monthStep || !state) return;

  const current = fromIso(state.viewMonth) || startOfMonth(startOfToday());
  let nextMonthStart;

  if (monthStep === 'today') {
    nextMonthStart = startOfMonth(startOfToday());
  } else {
    const offset = Number(monthStep);
    if (!Number.isFinite(offset)) {
      return;
    }
    nextMonthStart = addMonths(current, offset);
  }

  state.viewMonth = toIso(startOfMonth(nextMonthStart));
  saveState();
  renderMonthView({
    planReady: state.planCommitted,
    startDate: fromIso(state.startDate),
    deadline: resolveDeadline(state),
    today: startOfToday(),
    schedule: buildScheduleMap(),
  });
}

function syncDeadlineInput() {
  if (!state) return;
  const startValue = elements.startInput?.value;
  if (startValue && elements.endInput) {
    elements.endInput.min = startValue;
  }

  const radios = elements.timelineRadios();
  const selected = radios ? radios.value : '';

  if (selected && selected !== 'custom' && startValue) {
    const weeks = Number(selected);
    const computedEnd = toIso(addDays(fromIso(startValue), weeks * 7 - 1));
    if (elements.endInput) {
      elements.endInput.value = computedEnd;
      elements.endInput.disabled = true;
    }
  } else if (selected === 'custom') {
    if (elements.endInput) {
      elements.endInput.disabled = false;
      if (!elements.endInput.value && state.customEndDate) {
        elements.endInput.value = state.customEndDate;
      }
    }
  } else if (elements.endInput) {
    elements.endInput.disabled = false;
  }
}

function render() {
  const today = startOfToday();
  const todayKey = toIso(today);
  const planReady = Boolean(
    state &&
      state.planCommitted &&
      state.startDate &&
      (state.customEndDate || state.timelineWeeks),
  );

  updatePlanForm(planReady);
  renderPlanVisibility(planReady);

  const startDate = planReady ? fromIso(state.startDate) : null;
  const deadline = planReady ? resolveDeadline(state) : null;
  const totalTasks = state.tasks.length;
  const completedTasks = state.tasks.filter((task) => task.completed).length;
  const pendingTasks = totalTasks - completedTasks;
  const progress = totalTasks
    ? Math.round((completedTasks / totalTasks) * 100)
    : 0;

  updateStats({
    progress,
    completedTasks,
    totalTasks,
    pendingTasks,
    startDate,
    deadline,
    today,
  });

  let schedule = new Map();
  if (planReady && startDate && deadline) {
    const rebalanced = rebalanceOverdueTasks(today, deadline);
    if (rebalanced) {
      saveState();
    }
    schedule = buildScheduleMap();
  }

  const dayConfigs = [
    { key: 'yesterday', date: addDays(today, -1) },
    { key: 'today', date: today },
    { key: 'tomorrow', date: addDays(today, 1) },
  ];

  for (const config of dayConfigs) {
    renderDayCard({
      ...config,
      schedule,
      startDate,
      deadline,
      today,
      planReady,
      todayKey,
    });
  }

  renderMonthView({
    schedule,
    planReady,
    startDate,
    deadline,
    today,
  });
}

function updatePlanForm(planReady) {
  const radios = elements.timelineRadios();
  if (radios) {
    radios.value = state.timelineMode || '';
  }

  if (elements.startInput) {
    elements.startInput.value = state.startDate || '';
  }

  const summaryDays =
    state.studyDays && state.studyDays.length
      ? state.studyDays
      : defaultStudyDays();
  setStudyDaySelections(summaryDays);
  updateStudyDaySummary(summaryDays);

  if (state.timelineMode === 'custom' && state.customEndDate) {
    if (elements.endInput) {
      elements.endInput.value = state.customEndDate;
      elements.endInput.disabled = false;
    }
  } else if (planReady && state.timelineWeeks && state.startDate) {
    const computed = toIso(
      addDays(fromIso(state.startDate), state.timelineWeeks * 7 - 1),
    );
    if (elements.endInput) {
      elements.endInput.value = computed;
      if (state.timelineMode !== 'custom') {
        elements.endInput.disabled = true;
      }
    }
  } else if (!planReady && elements.endInput) {
    elements.endInput.value = '';
    elements.endInput.disabled = state.timelineMode !== 'custom';
  }

  syncDeadlineInput();
  clearStudyDayError();
}

function renderPlanVisibility(planReady) {
  const config = elements.plannerConfig;
  if (!config) return;
  const shouldShow = !planReady || isEditingPlan;
  config.classList.toggle('planner-config--hidden', !shouldShow);
  config.setAttribute('aria-hidden', String(!shouldShow));

  if (elements.editButton) {
    elements.editButton.style.display = planReady ? 'inline' : 'none';
    elements.editButton.setAttribute('aria-expanded', String(shouldShow));
    elements.editButton.title = planReady
      ? `Study days: ${formatStudyDaySummary(state.studyDays || defaultStudyDays())}`
      : 'Edit plan';
  }

  if (elements.exportButton) {
    const canExport = planReady && !isEditingPlan;
    elements.exportButton.style.display = canExport ? 'inline' : 'none';
    elements.exportButton.title = planReady
      ? canExport
        ? 'Download a PDF copy of your plan'
        : 'Finish editing to export your plan'
      : 'Save your plan to enable exporting';
  }

  if (elements.resetButton) {
    const hasProgress = state.tasks && state.tasks.some((task) => task.completed);
    elements.resetButton.style.display = hasProgress ? 'inline' : 'none';
  }
}

function updateStats({
  progress,
  completedTasks,
  totalTasks,
  pendingTasks,
  startDate,
  deadline,
  today,
}) {
  elements.stats.progressValue.textContent = `${progress}%`;
  elements.stats.progressDetail.textContent = `${completedTasks} of ${totalTasks} tasks done`;
  elements.stats.progressBar.style.width = `${progress}%`;

  elements.stats.tasksRemaining.textContent = pendingTasks;

  if (!startDate || !deadline) {
    elements.stats.tasksPerDayHint.textContent = '0 per day';
    elements.stats.daysRemaining.textContent = '0';
    elements.stats.deadlineLabel.textContent = 'Deadline — Set your plan';
    return;
  }

  const clampedToday =
    today < startDate ? startDate : today > deadline ? deadline : today;
  const remainingDays = Math.max(diffDays(clampedToday, deadline) + 1, 0);
  const allowedDays = getAllowedDaysSet();
  const workingDaysRemaining =
    remainingDays > 0
      ? countAllowedDaysBetween(clampedToday, deadline, allowedDays)
      : 0;
  const tasksPerDay =
    workingDaysRemaining > 0
      ? Math.max(Math.ceil(pendingTasks / workingDaysRemaining), 0)
      : pendingTasks;

  elements.stats.tasksPerDayHint.textContent =
    workingDaysRemaining > 0
      ? `${tasksPerDay} per study day`
      : pendingTasks > 0
      ? 'No study days remaining'
      : 'All done';

  const daysLeft = today > deadline ? 0 : Math.max(diffDays(today, deadline) + 1, 0);
  elements.stats.daysRemaining.textContent = String(daysLeft);
  elements.stats.deadlineLabel.textContent = `Deadline — ${
    deadline ? formatLongDate(deadline) : 'Set your plan'
  }`;
}

function renderDayCard({
  key,
  date,
  schedule,
  startDate,
  deadline,
  today,
  planReady,
  todayKey,
}) {
  const card = elements.cards[key];
  if (!card) return;

  const dateKey = toIso(date);
  card.classList.remove('day-card--drop-target');
  card.dataset.dateKey = dateKey;
  const label = card.querySelector('.day-label');
  if (label && key === 'today') {
    label.textContent = date.getTime() === today.getTime() ? 'Today' : 'Today*';
  }
  const dateEl = card.querySelector('.day-date');
  if (dateEl) {
    dateEl.textContent = formatShortDate(date);
  }

  const list = card.querySelector('.task-list');
  list.innerHTML = '';

  const emptyState = card.querySelector('.empty-state');

  if (!planReady) {
    setEmptyState(emptyState, 'Choose a start date and timeline to generate your plan.');
    card.classList.add('day-card--inactive');
    return;
  }

  card.classList.remove('day-card--inactive');

  const isBeforePlan = startDate && date < startDate;
  const isAfterPlan = deadline && date > deadline;
  const isPlanDay =
    startDate && deadline && date >= startDate && date <= deadline;

  let tasksForDay = [];
  let infoMessage = '';

  if (isBeforePlan) {
    infoMessage = `Plan starts on ${formatShortDate(startDate)}.`;
  } else if (isAfterPlan) {
    const overdue = state.tasks.filter((task) => !task.completed);
    if (overdue.length) {
      infoMessage = `Deadline passed on ${formatShortDate(deadline)} — ${overdue.length} tasks overdue.`;
      tasksForDay = overdue;
    } else {
      infoMessage = 'Plan complete. Great job!';
    }
  } else if (isPlanDay && schedule.has(dateKey)) {
    tasksForDay = [...(schedule.get(dateKey) || [])];
  }

  const completedToday = state.tasks.filter(
    (task) => task.completed && task.completedOn === dateKey,
  );

  const combined = [...tasksForDay, ...completedToday];
  const uniqueTasks = new Map();
  combined.forEach((task) => {
    uniqueTasks.set(task.id, task);
  });

  const sortedTasks = Array.from(uniqueTasks.values()).sort((a, b) => {
    if (a.completed === b.completed) {
      return a.order - b.order;
    }
    return a.completed ? 1 : -1;
  });

  if (!sortedTasks.length) {
    if (infoMessage) {
      setEmptyState(emptyState, infoMessage);
    } else {
      const message =
        dateKey === todayKey
          ? 'Everything done for today.'
          : key === 'tomorrow'
          ? 'No tasks scheduled. Consider reviewing past notes.'
          : 'No tasks were scheduled.';
      setEmptyState(emptyState, message);
    }
    return;
  }

  if (infoMessage) {
    setEmptyState(emptyState, infoMessage);
  } else if (emptyState) {
    emptyState.textContent = '';
  }

  sortedTasks.forEach((task) => {
    const node = createTaskNode(task, dateKey);
    if (isAfterPlan && !task.completed) {
      node.classList.add('task-item--overdue');
    }
    list.appendChild(node);
  });
}

function createTaskNode(task, dateKey) {
  const fragment = elements.template.content.cloneNode(true);
  const item = fragment.querySelector('.task-item');
  const checkbox = fragment.querySelector('input[type="checkbox"]');
  const title = fragment.querySelector('.task-title');
  const detail = fragment.querySelector('.task-detail');

  item.dataset.taskId = task.id;
  item.dataset.dateKey = dateKey;

  const isComplete = Boolean(task.completed);
  checkbox.checked = isComplete;
  checkbox.disabled = false;

  item.draggable = !isComplete;
  item.classList.toggle('task-item--draggable', !isComplete);

  if (isComplete) {
    item.classList.add('task-item--complete');
    item.draggable = false;
  } else {
    item.classList.remove('task-item--complete');
  }

  title.textContent = task.title;
  const skillCopy = task.skill ? task.skill.replace(/^Skill\s+\d+\s+-\s+/i, '') : '';
  const parts = [];
  if (skillCopy) parts.push(skillCopy);
  if (task.type) parts.push(task.type);
  detail.textContent = parts.join(' • ');

  return fragment;
}

function rebuildScheduleAssignments(startDate, deadline) {
  if (!startDate || !deadline || !state) return;
  const sortedTasks = [...state.tasks].sort((a, b) => a.order - b.order);
  const totalDays = diffDays(startDate, deadline) + 1;
  const allowedDays = getAllowedDaysSet();

  if (totalDays <= 0) {
    const singleDay = toIso(startDate);
    sortedTasks.forEach((task) => {
      task.scheduledDate = singleDay;
    });
    return;
  }

  const dayKeys = [];
  for (let offset = 0; offset < totalDays; offset++) {
    const current = addDays(startDate, offset);
    if (allowedDays.has(current.getDay())) {
      dayKeys.push(toIso(current));
    }
  }

  if (!dayKeys.length) {
    dayKeys.push(toIso(deadline));
  }

  sortedTasks.forEach((task) => {
    task.scheduledDate = null;
  });

  let pointer = 0;
  for (let index = 0; index < dayKeys.length; index++) {
    const dayKey = dayKeys[index];
    const daysRemaining = dayKeys.length - index;
    const tasksRemaining = sortedTasks.length - pointer;
    const take =
      daysRemaining <= 0 || tasksRemaining <= 0
        ? 0
        : Math.ceil(tasksRemaining / daysRemaining);

    for (let count = 0; count < take && pointer < sortedTasks.length; count++) {
      sortedTasks[pointer++].scheduledDate = dayKey;
    }
  }

  const fallbackDay = dayKeys[dayKeys.length - 1];
  while (pointer < sortedTasks.length) {
    sortedTasks[pointer++].scheduledDate = fallbackDay;
  }
}

function rebalanceOverdueTasks(today, deadline) {
  if (deadline && deadline < today) {
    return false;
  }
  const todayKey = toIso(today);
  const allowedDays = getAllowedDaysSet();
  const overdueTasks = state.tasks.filter(
    (task) =>
      !task.completed &&
      task.scheduledDate &&
      task.scheduledDate < todayKey,
  );
  if (!overdueTasks.length) {
    return false;
  }

  const overdueSet = new Set(overdueTasks.map((task) => task.id));
  overdueTasks.sort((a, b) => a.order - b.order);

  const dayKeys = [];
  const span = Math.max(diffDays(today, deadline) + 1, 0);
  if (span <= 0) {
    if (allowedDays.has(today.getDay())) {
      dayKeys.push(todayKey);
    }
    if (!dayKeys.length) {
      dayKeys.push(toIso(deadline));
    }
  } else {
    for (let offset = 0; offset < span; offset++) {
      const current = addDays(today, offset);
      if (allowedDays.has(current.getDay())) {
        dayKeys.push(toIso(current));
      }
    }
  }

  if (!dayKeys.length) {
    dayKeys.push(toIso(deadline));
  }

  const loadMap = new Map();
  dayKeys.forEach((key) => {
    const existing = state.tasks.filter(
      (task) =>
        !task.completed &&
        task.scheduledDate === key &&
        !overdueSet.has(task.id),
    );
    loadMap.set(key, existing.length);
  });

  overdueTasks.forEach((task) => {
    task.scheduledDate = null;
  });

  overdueTasks.forEach((task) => {
    let targetKey = dayKeys[0];
    let smallestLoad = Number.POSITIVE_INFINITY;
    for (const key of dayKeys) {
      const load = loadMap.get(key) ?? 0;
      if (load < smallestLoad) {
        smallestLoad = load;
        targetKey = key;
      }
    }
    loadMap.set(targetKey, (loadMap.get(targetKey) ?? 0) + 1);
    task.scheduledDate = targetKey;
  });

  return true;
}

function buildScheduleMap() {
  const schedule = new Map();
  state.tasks.forEach((task) => {
    if (!task.scheduledDate) return;
    if (!schedule.has(task.scheduledDate)) {
      schedule.set(task.scheduledDate, []);
    }
    schedule.get(task.scheduledDate).push(task);
  });

  schedule.forEach((tasks, key) => {
    tasks.sort((a, b) => a.order - b.order);
    schedule.set(key, tasks);
  });

  return schedule;
}

function renderMonthView({ schedule, planReady, startDate, deadline, today }) {
  if (!elements.month.grid || !elements.month.label) return;

  const monthDate = fromIso(state.viewMonth) || startOfMonth(today);
  const normalizedMonth = startOfMonth(monthDate);
  state.viewMonth = toIso(normalizedMonth);

  elements.month.label.textContent = formatMonthYear(normalizedMonth);

  const monthView = document.getElementById('monthView');
  if (!planReady || !startDate || !deadline) {
    if (monthView) {
      monthView.classList.add('is-locked');
    }
    elements.month.grid.innerHTML = '';
    return;
  }

  if (monthView) {
    monthView.classList.remove('is-locked');
  }

  const calendarStart = startOfWeek(normalizedMonth);
  const calendarEnd = endOfWeek(endOfMonth(normalizedMonth));
  const fragment = document.createDocumentFragment();
  const todayTime = today.getTime();

  for (
    let cursor = new Date(calendarStart);
    cursor <= calendarEnd;
    cursor = addDays(cursor, 1)
  ) {
    const cell = document.createElement('div');
    cell.className = 'calendar-cell';
    const dateKey = toIso(cursor);
    const dayNumber = cursor.getDate();
    const isCurrentMonth = cursor.getMonth() === normalizedMonth.getMonth();
    const isToday = cursor.getTime() === todayTime;
    const tasksForDay = schedule.get(dateKey) || [];
    const dueCount = tasksForDay.filter((task) => !task.completed).length;
    const doneCount = tasksForDay.length - dueCount;
    const isPast = cursor < today;

    if (!isCurrentMonth) {
      cell.classList.add('calendar-cell--outside');
    }
    if (isToday) {
      cell.classList.add('calendar-cell--today');
    }
    if (dueCount > 0 && isPast) {
      cell.classList.add('calendar-cell--overdue');
    }
    if (dueCount === 0 && doneCount > 0) {
      cell.classList.add('calendar-cell--complete');
    }
    if (dueCount > 0 && !isPast) {
      cell.classList.add('calendar-cell--upcoming');
    }

    const dateLabel = document.createElement('span');
    dateLabel.className = 'calendar-day';
    dateLabel.textContent = String(dayNumber);
    cell.appendChild(dateLabel);

    if (tasksForDay.length) {
      const badges = document.createElement('div');
      badges.className = 'calendar-badges';
      if (dueCount > 0) {
        const dueBadge = document.createElement('span');
        dueBadge.className = 'calendar-badge calendar-badge--pending';
        dueBadge.textContent = `${dueCount} due`;
        badges.appendChild(dueBadge);
      }
      if (doneCount > 0) {
        const doneBadge = document.createElement('span');
        doneBadge.className = 'calendar-badge calendar-badge--done';
        doneBadge.textContent = `${doneCount} done`;
        badges.appendChild(doneBadge);
      }
      cell.appendChild(badges);
    }

    const ariaParts = [`${formatShortDate(cursor)}`];
    if (dueCount > 0) ariaParts.push(`${dueCount} due`);
    if (doneCount > 0) ariaParts.push(`${doneCount} done`);
    cell.setAttribute('aria-label', ariaParts.join(', '));

    fragment.appendChild(cell);
  }

  elements.month.grid.innerHTML = '';
  elements.month.grid.appendChild(fragment);
}

function handleDragStart(event) {
  const item = event.target.closest('.task-item');
  if (!item || item.classList.contains('task-item--complete')) {
    return;
  }
  if (!state?.planCommitted) {
    return;
  }
  const taskId = item.dataset.taskId;
  const dateKey = item.dataset.dateKey;
  if (!taskId || !dateKey) {
    return;
  }
  dragState = { taskId, fromDate: dateKey };
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', taskId);
  }
  item.classList.add('task-item--dragging');
}

function handleDragEnd(event) {
  const item = event.target.closest('.task-item');
  item?.classList.remove('task-item--dragging');
  cleanupDropTargets();
  dragState = null;
}

function handleDragEnter(event) {
  if (!dragState) return;
  const card = event.target.closest('.day-card');
  if (!card) return;
  card.classList.add('day-card--drop-target');
}

function handleDragOver(event) {
  if (!dragState) return;
  const card = event.target.closest('.day-card');
  if (!card) return;
  event.preventDefault();
  if (event.dataTransfer) {
    event.dataTransfer.dropEffect = 'move';
  }
  card.classList.add('day-card--drop-target');
}

function handleDragLeave(event) {
  if (!dragState) return;
  const card = event.target.closest('.day-card');
  if (!card) return;
  const related = event.relatedTarget;
  if (!related || !card.contains(related)) {
    card.classList.remove('day-card--drop-target');
  }
}

function handleDrop(event) {
  if (!dragState || !state?.planCommitted) return;
  const card = event.target.closest('.day-card');
  if (!card) return;
  event.preventDefault();

  const targetDate = card.dataset.dateKey;
  if (!targetDate) return;

  const task = state.tasks.find((item) => item.id === dragState.taskId);
  if (!task || task.completed) {
    cleanupDropTargets();
    dragState = null;
    return;
  }

  if (task.scheduledDate !== targetDate) {
    task.scheduledDate = targetDate;
    task.completedOn = null;
    saveState();
  }

  cleanupDropTargets();
  dragState = null;
  render();
}

function cleanupDropTargets() {
  Object.values(elements.cards).forEach((card) =>
    card?.classList.remove('day-card--drop-target'),
  );
  document
    .querySelectorAll('.task-item--dragging')
    .forEach((node) => node.classList.remove('task-item--dragging'));
}

function getSelectedStudyDays() {
  const inputs = elements.studyDayInputs();
  if (!inputs) return [];
  const selection = new Set();
  Array.from(inputs).forEach((input) => {
    if (input.checked) {
      const num = Number(input.value);
      if (Number.isInteger(num) && num >= 0 && num <= 6) {
        selection.add(num);
      }
    }
  });
  return Array.from(selection).sort((a, b) => a - b);
}

function setStudyDaySelections(days) {
  const inputs = elements.studyDayInputs();
  if (!inputs) return;
  const target = new Set(
    normaliseDayList(days, { fallback: true }).map((value) => Number(value)),
  );
  Array.from(inputs).forEach((input) => {
    const num = Number(input.value);
    input.checked = target.has(num);
  });
}

function updateStudyDaySummary(days) {
  const node = elements.studyDaySummary;
  if (!node) return;
  const list = normaliseDayList(days, { fallback: false });
  if (!list.length) {
    node.textContent = 'No study days selected yet.';
    return;
  }
  if (list.length === 7) {
    node.textContent = 'Selected: Every day';
    return;
  }
  node.textContent = `Selected: ${list
    .map((day) => DAY_LABELS_SHORT[day])
    .join(' · ')}`;
}

function showStudyDayError(message) {
  if (elements.studyDayFieldset) {
    elements.studyDayFieldset.classList.add('weekdays-fieldset--error');
  }
  if (elements.studyDayError) {
    elements.studyDayError.textContent = message;
  }
}

function clearStudyDayError() {
  elements.studyDayFieldset?.classList.remove('weekdays-fieldset--error');
  if (elements.studyDayError) {
    elements.studyDayError.textContent = '';
  }
}

function getAllowedDaysSet() {
  const list = normaliseDayList(state?.studyDays, { fallback: true });
  return new Set(list);
}

function defaultStudyDays() {
  return [1, 2, 3, 4, 5];
}

function sanitiseStudyDays(days) {
  return normaliseDayList(days, { fallback: true });
}

function normaliseDayList(days, { fallback = false } = {}) {
  const unique = new Set();
  if (Array.isArray(days)) {
    days.forEach((value) => {
      const num = Number(value);
      if (Number.isInteger(num) && num >= 0 && num <= 6) {
        unique.add(num);
      }
    });
  }
  if (!unique.size && fallback) {
    defaultStudyDays().forEach((value) => unique.add(value));
  }
  return Array.from(unique).sort((a, b) => a - b);
}

function formatStudyDaySummary(days) {
  const list = normaliseDayList(days, { fallback: false });
  if (!list.length) {
    return 'No study days selected';
  }
  if (list.length === 7) {
    return 'Every day';
  }
  return list.map((day) => DAY_LABELS_LONG[day]).join(', ');
}

function countAllowedDaysBetween(startDate, endDate, allowedDays) {
  if (!startDate || !endDate) return 0;
  let count = 0;
  for (
    let cursor = new Date(startDate);
    cursor <= endDate;
    cursor = addDays(cursor, 1)
  ) {
    if (allowedDays.has(cursor.getDay())) {
      count += 1;
    }
  }
  return count;
}

function startOfToday() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function addMonths(date, months) {
  const result = new Date(date);
  result.setMonth(result.getMonth() + months, 1);
  return startOfMonth(result);
}

function diffDays(from, to) {
  const startUTC = Date.UTC(from.getFullYear(), from.getMonth(), from.getDate());
  const endUTC = Date.UTC(to.getFullYear(), to.getMonth(), to.getDate());
  return Math.round((endUTC - startUTC) / DAY_MS);
}

function toIso(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function fromIso(iso) {
  if (!iso) return null;
  const [year, month, day] = iso.split('-').map(Number);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return null;
  }
  return new Date(year, month - 1, day);
}

function resolveDeadline(localState) {
  if (localState.timelineMode === 'custom' && localState.customEndDate) {
    return fromIso(localState.customEndDate);
  }
  if (localState.timelineWeeks && localState.startDate) {
    return addDays(fromIso(localState.startDate), localState.timelineWeeks * 7 - 1);
  }
  return null;
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function endOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function startOfWeek(date) {
  const result = new Date(date);
  const day = result.getDay();
  return addDays(result, -day);
}

function endOfWeek(date) {
  const result = new Date(date);
  const day = result.getDay();
  return addDays(result, 6 - day);
}

function formatShortDate(date) {
  return new Intl.DateTimeFormat(undefined, {
    weekday: 'short',
    month: 'short',
    day: 'numeric',
  }).format(date);
}

function formatLongDate(date) {
  return new Intl.DateTimeFormat(undefined, {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  }).format(date);
}

function formatMonthYear(date) {
  return new Intl.DateTimeFormat(undefined, {
    month: 'long',
    year: 'numeric',
  }).format(date);
}

function capitalise(value) {
  if (!value) return value;
  return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
}

function compareIsoDates(a, b) {
  return new Date(a) - new Date(b);
}

function setEmptyState(node, message) {
  if (!node) return;
  node.textContent = message;
}

function displayCourseLoadError(error) {
  const todayCard = elements.cards.today;
  const message = error
    ? 'Unable to load the course outline. Refresh or check the file path.'
    : 'Unable to load the course outline.';
  const emptyState = todayCard?.querySelector('.empty-state');
  setEmptyState(emptyState, message);
}

function toRadioArray(radioNodeList) {
  if (!radioNodeList) return [];
  if (typeof radioNodeList.length === 'number') {
    return Array.from(radioNodeList);
  }
  return [radioNodeList];
}
