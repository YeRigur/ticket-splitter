const escapeRegex = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export const normalizeProjects = (projects) => {
  if (!projects) {
    return [];
  }
  return (Array.isArray(projects) ? projects : [projects])
    .map((project) => String(project).trim())
    .filter(Boolean);
};

export const normalizeWords = (values) => {
  if (!values) {
    return [];
  }
  return (Array.isArray(values) ? values : [values])
    .map((value) => String(value).trim())
    .filter(Boolean);
};

export const normalizeText = (value) =>
  String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[_/|\\()[\]{}.,;:!?'"`-]+/g, ' ')
    .toLowerCase();

export const looksLikeRegex = (value) => {
  if (typeof value !== 'string') {
    return false;
  }
  return (
    value.startsWith('\\b') ||
    (value.includes('\\.') && value.includes('.*')) ||
    /[\[\]{}()+?^$|]/.test(value)
  );
};

export const legacyPatternToMatcher = (pattern) => {
  const value = String(pattern ?? '').trim();
  if (!value) {
    return null;
  }

  if (looksLikeRegex(value)) {
    return { type: 'regex', value };
  }

  if (value.includes('*')) {
    const pieces = value
      .split('*')
      .map((piece) => piece.trim())
      .filter(Boolean);
    if (pieces.length > 1) {
      return { type: 'containsAll', values: pieces };
    }
    if (pieces.length === 1) {
      return { type: 'contains', value: pieces[0] };
    }
    return null;
  }

  return { type: 'contains', value };
};

export const normalizeMatcher = (matcher, { context = 'project' } = {}) => {
  if (!matcher || typeof matcher !== 'object') {
    throw new Error(`Invalid ${context} matcher.`);
  }

  const type = String(matcher.type ?? 'contains').trim();
  if (type === 'contains') {
    const value = String(matcher.value ?? '').trim();
    return { type, value };
  }

  if (type === 'containsWord') {
    const value = String(matcher.value ?? '').trim();
    return { type, value };
  }

  if (type === 'containsAny' || type === 'containsAll') {
    const values = normalizeWords(matcher.values);
    return { type, values };
  }

  if (type === 'regex') {
    const value = String(matcher.value ?? '').trim();
    return { type, value };
  }

  throw new Error(`Unsupported ${context} matcher type "${type}".`);
};

export const normalizeDefinitions = (definitions, { scoped = false, context = 'project' } = {}) => {
  if (!Array.isArray(definitions)) {
    throw new Error(`${context} config must be an array.`);
  }

  return definitions.map((entry, index) => {
    if (!entry || typeof entry !== 'object') {
      throw new Error(`Rule at index ${index} must be an object.`);
    }

    const name = String(entry.name ?? '').trim();
    if (!name) {
      throw new Error(`Rule at index ${index} is missing a name.`);
    }

    const rules = Array.isArray(entry.rules)
      ? entry.rules.map((rule) => normalizeMatcher(rule, { context: `${context} rule "${name}"` }))
      : Array.isArray(entry.patterns)
        ? entry.patterns
            .map(legacyPatternToMatcher)
            .filter(Boolean)
        : [];

    if (!rules.length) {
      throw new Error(`Rule "${name}" must include at least one matcher.`);
    }

    const normalized = { name, rules };
    if (scoped) {
      normalized.projects = normalizeProjects(entry.projects);
    }
    return normalized;
  });
};

export const legacyDefinitionsToStructured = (definitions, { scoped = false, context = 'project' } = {}) =>
  normalizeDefinitions(definitions, { scoped, context });

const matcherMatches = (matcher, originalText, normalizedText) => {
  if (matcher.type === 'contains') {
    const value = normalizeText(matcher.value);
    return value ? normalizedText.includes(value) : false;
  }
  if (matcher.type === 'containsWord') {
    const normalizedValue = normalizeText(matcher.value);
    if (!normalizedValue) {
      return false;
    }
    const boundaryRegex = new RegExp(`(^|\\s)${escapeRegex(normalizedValue)}(\\s|$)`, 'i');
    return boundaryRegex.test(normalizedText);
  }
  if (matcher.type === 'containsAny') {
    return normalizeWords(matcher.values).some((value) => normalizedText.includes(normalizeText(value)));
  }
  if (matcher.type === 'containsAll') {
    const values = normalizeWords(matcher.values);
    return values.length ? values.every((value) => normalizedText.includes(normalizeText(value))) : false;
  }
  if (matcher.type === 'regex') {
    if (!String(matcher.value ?? '').trim()) {
      return false;
    }
    try {
      const regex = new RegExp(matcher.value, 'i');
      return regex.test(originalText) || regex.test(normalizedText);
    } catch (error) {
      console.warn(`Invalid regex matcher "${matcher.value}" ignored.`, error);
      return false;
    }
  }
  return false;
};

export const compileDefinitions = (definitions, { scoped = false } = {}) =>
  normalizeDefinitions(definitions, { scoped }).map((definition) => ({
    name: definition.name,
    projects: scoped ? definition.projects : [],
    matchers: definition.rules
  }));

export const matchesDefinition = (definition, text) => {
  const originalText = String(text ?? '');
  const normalizedText = normalizeText(originalText);
  return definition.matchers.some((matcher) => matcherMatches(matcher, originalText, normalizedText));
};
