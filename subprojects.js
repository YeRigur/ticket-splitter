import { compileDefinitions, normalizeDefinitions } from './rule-engine.js';

export const normalizeSubProjectDefinitions = (definitions) =>
  normalizeDefinitions(definitions, { scoped: true, context: 'subproject' });

export const compileSubProjectFilters = (definitions) =>
  compileDefinitions(definitions, { scoped: true });
