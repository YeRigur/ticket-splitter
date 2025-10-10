const filters = [
  {
    name: 'Technical & Content',
    patterns: [
      'NBSTTTPGL*',
      'TTT PGL Alignment',
      'PGL Alignment Tuesday',
      'PGL Alignment Thursday',
      'DAILY | PGL | Content&QA + Tech',
      'PGL Alignment call',
      'PGL Alignment',
      'PGL requests management',
      'TTT PGL Changes for Brands banner component*',
      '== PGL Alignment ==',
      '\\bpgl\\b',
      'NBSLK*',
      'NBS tech x Purina UK weekly*',
      'Alignment Call 2 - UX/UI + Tech',
      '*POSTDEPLOYFIXES*',
      'Gigya IT Suppurt',
      'TINT API maintenance*',
      'Pagination indexing topic discussion',
      'Newsletter In-Context - Experian validation',
      'Project set up',
      'Project estimation alignment',
      'Project Details Sharing',
      'Project Tracker Review',
      'Experiment Tracker Alignement',
      'Data catch up',
      'PLP data alignment',
      'Local Deploy*',
      'Deploy local',
      'Drusher for Pantheon links & DBs',
      'CI registration',
      'Catch-up*',
      'ZT Team weekly',
      'Purina weekly cooridnators',
      'Website Services Team Leadership Announcement',
      'Intro call',
      'tickets creation',
      'Creating tickets for Sprint',
      'Ticket management',
      'Task consultation',
      'Ticket internal discussion',
      'tickets reviewing',
      'EMENA;GLOBAL;PUR;ALL;Non-ticket*',
      'Non*ticket*',
      'UK*non-ticket*',
      'R&R not being displayed at products',
      '*Tint*',
      '*Scroll*',
      'Urgent requests statistics*',
      'mandatory trainings',
      'Reading emails',
      'Help for a colleague*',
      'UK Weekly',
      'UK*activities',
      'UK*activity',
      'UK Consumer webform',
      'UK Contact Us Webforms',
      'UK*tickets*',
      'UK support',
      '*Salesforce Mapping UK*',
      'Feeding Tool UK*',
      'Front-end coding',
      'Adimo*',
      'TASKINVESTIGATION'
    ]
  },
  {
    name: 'Light House',
    patterns: [
      'NBSTTTME*',
      '*NBSTTTME*',
      'NBSTTME*',
      '*NBSTTME*',
      'ME Weekly*',
      'ME calculation new prototype',
      'ME*phase planning',
      'Me/UK deploying*',
      '.*[Ff][Oo][Oo][Dd].*[Rr][Ee][Cc][Oo].*',
      '.*[Ff][Oo][Oo][Dd][Rr][Ee][Cc][Oo].*',
      'FoodReco*',
      'Food Reco*',
      'food reco*',
      'FOOD RECO*',
      '*FoodReco*',
      '*Food Reco*',
      '*food reco*',
      '*FOOD RECO*',
      'FoodReco status align',
      'Food Reco Catch Up',
      'Food recco*',
      'FRT*',
      '*FRT*',
      'FRT Biweekly*',
      '*Feeding.*tool*',
      'Feeding Tool*',
      'Purina Feeding Guide*',
      'Naming Tool*'
    ]
  },
  {
    name: 'Web Enhancements & Braze',
    patterns: [
      '*NBSTTTE*',
      '*Enhancement*',
      'Enhancement*',
      'Sprint events - enhancement squad',
      'Daily standup - enhancement squad',
      'Enhancements*',
      'Navigation Menu*',
      'Catch-up*Navigation*',
      '\\bAB\\b',
      '\\bA/B\\b',
      '*VWO*',
      'A/B testing*',
      'Homepage optimization*',
      'Homepage*design*',
      '*Braze*',
      '*braze*',
      'Inbox Notification*',
      'Content cards*',
      'Article*',
      'Articles*',
      '*Accessibility*',
      '\\bAC\\b',
      'A11y',
      '*Accessibilit*',
      'Nestle and Deque*',
      '*seo*',
      'Defer offscreen images*',
      'Reduce unused CSS*',
      'Minimize main-thread work',
      'Changes for Brands banner component',
      'Felix Brand hub banner*',
      '*Bazaarvoice*',
      '\\[[A-Z]{2}\\].*',
      'Sprint Review*',
      'Sprint Retro*',
      'FAQ Question scheme*',
      'Breed Selector*',
      'Typographic scale*',
      'purina\\..*— Other Ticket*',
      'purina\\..*— Color Ticket*',
      'Ticket Validation Flow',
      'Purina\\.com Product Update*',
      'Projects alignment',
      'Roadmap opportunities'
    ]
  },
  {
    name: 'Project communication',
    patterns: [
      'Purina coordinators*',
      'Purina.*coordinators*',
      'Purina Europe.*alignment*',
      '*Project Alignment*',
      '*RZ*',
      'LATAM.*Interoperability.*Catch-up*'
    ]
  },
  {
    name: 'Incorrect codes',
    patterns: [
      'Tech.*team.*call*',
      'Tech.*Team.*WOR*',
      'Tech.*Team.*Weekly*',
      '*Drupal Club*',
      '\\bPI\\b',
      '*Personal*',
      '\\bAOA\\b',
      'March IMS Floor Walk*',
      'Automation Express webinar*',
      'Green forest English course*',
      'Purina Institute.*alignment*',
      'Food Code Corrections*'
    ]
  }
];

const escapeRegex = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const toRegex = (pattern) => {
  if (pattern.startsWith('\\b') || (pattern.includes('\\.') && pattern.includes('.*'))) {
    return new RegExp(pattern, 'i');
  }
  if (pattern.includes('*')) {
    const cleaned = pattern.replace(/^\*+|\*+$/g, '');
    if (!cleaned) {
      return new RegExp('.*', 'i');
    }
    if (pattern.startsWith('*') && pattern.endsWith('*')) {
      return new RegExp(`.*${escapeRegex(cleaned)}.*`, 'i');
    }
    if (pattern.startsWith('*')) {
      return new RegExp(`.*${escapeRegex(cleaned)}`, 'i');
    }
    if (pattern.endsWith('*')) {
      return new RegExp(`${escapeRegex(cleaned)}.*`, 'i');
    }
    const escaped = escapeRegex(pattern).replace(/\\\*/g, '.*');
    return new RegExp(escaped, 'i');
  }
  return new RegExp(escapeRegex(pattern), 'i');
};

export const compileFilters = (definitions) =>
  definitions.map(({ name, patterns }) => {
    const regexes = [];
    for (const pattern of patterns) {
      try {
        regexes.push(toRegex(pattern));
      } catch (error) {
        console.warn(`Skipping pattern "${pattern}" for project "${name}":`, error);
      }
    }
    return { name, regexes };
  });

export const filterDefinitions = filters;
export const realProjects = filters
  .filter((entry) => entry.name !== 'Incorrect codes')
  .map((entry) => entry.name);
