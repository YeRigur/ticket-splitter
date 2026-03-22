import { filterDefinitions } from './filters.js';
import { legacyDefinitionsToStructured } from './rule-engine.js';

export const defaultProjectDefinitions = legacyDefinitionsToStructured(filterDefinitions, {
  scoped: false,
  context: 'project'
});
