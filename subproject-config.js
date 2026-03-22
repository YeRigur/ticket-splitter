import { legacyDefinitionsToStructured } from './rule-engine.js';

const subProjectDefinitionsRaw = [
  {
    projects: ['Technical & Content'],
    name: 'PGL',
    patterns: ['*PGL*']
  },
  {
    projects: ['Technical & Content'],
    name: 'NBS',
    patterns: ['*NBS*']
  },
  {
    projects: ['Technical & Content'],
    name: 'Deploy',
    patterns: ['*POSTDEPLOYFIXES*', 'Local Deploy*', 'Deploy local', 'Drusher for Pantheon links & DBs']
  },
  {
    projects: ['Light House'],
    name: 'Food Reco',
    patterns: ['*FoodReco*', '*Food Reco*', '*food reco*', '*FOOD RECO*']
  },
  {
    projects: ['Light House'],
    name: 'FRT',
    patterns: ['*FRT*']
  },
  {
    projects: ['Light House'],
    name: 'Feeding Tool',
    patterns: ['*Feeding Tool*', '*feeding tool*']
  },
  {
    projects: ['Light House'],
    name: 'Naming Tool',
    patterns: ['*Naming Tool*', '*naming tool*']
  },
  {
    projects: ['Light House'],
    name: 'ME',
    patterns: ['*NBSTTTME*', '*NBSTTME*', '\\bME\\b']
  },
  {
    projects: ['Web Enhancements & Braze'],
    name: 'Braze',
    patterns: ['*Braze*', '*braze*']
  },
  {
    projects: ['Web Enhancements & Braze'],
    name: 'Accessibility',
    patterns: ['*Accessibility*', '*Accessibilit*', '*A11y*']
  },
  {
    projects: ['Web Enhancements & Braze'],
    name: 'SEO',
    patterns: ['\\bseo\\b']
  },
  {
    projects: ['Web Enhancements & Braze'],
    name: 'Bazaarvoice',
    patterns: ['*Bazaarvoice*']
  },
  {
    projects: ['Web Enhancements & Braze'],
    name: 'Enhancement',
    patterns: ['*Enhancement*', '*Enhancements*']
  },
  {
    projects: ['Project communication'],
    name: 'Project Alignment',
    patterns: ['*Project Alignment*']
  },
  {
    projects: ['Project communication'],
    name: 'RZ',
    patterns: ['\\bRZ\\b']
  }
];

export const defaultSubProjectDefinitions = legacyDefinitionsToStructured(subProjectDefinitionsRaw, {
  scoped: true,
  context: 'subproject'
});
