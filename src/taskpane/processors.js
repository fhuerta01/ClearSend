/**
 * ClearSend Client-Side Processing Library
 *
 * PRIVACY GUARANTEE: All email recipient processing logic runs 100% locally.
 *
 * - Your email addresses NEVER leave your device
 * - All operations execute in your browser's memory
 * - No network calls transmit any email data
 * - No external servers process your recipient lists
 *
 * This file contains pure JavaScript functions that process email recipients
 * entirely within your Outlook application (desktop or web browser).
 */

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function extractEmail(recipient) {
  if (!recipient || typeof recipient !== 'string') {
    return '';
  }
  const match = recipient.match(/<(.+)>$/);
  return match ? match[1].trim() : recipient.trim();
}

function extractDisplayName(recipient) {
  if (!recipient || typeof recipient !== 'string') {
    return '';
  }
  const match = recipient.match(/^(.+?)\s*<(.+)>$/);
  if (match) {
    return match[1].trim();
  }
  return recipient.trim();
}

// ============================================================================
// SORT FUNCTIONS
// ============================================================================

function sortAlphabetical(recipients) {
  return recipients.sort((a, b) => {
    const nameA = extractDisplayName(a).toLowerCase();
    const nameB = extractDisplayName(b).toLowerCase();

    if (!nameA && !nameB) {
      const emailA = extractEmail(a).toLowerCase();
      const emailB = extractEmail(b).toLowerCase();
      return emailA.localeCompare(emailB);
    }

    return nameA.localeCompare(nameB);
  });
}

function sortStep(state) {
  const input = {
    to: [...state.to],
    cc: [...state.cc],
    bcc: [...state.bcc]
  };

  const sortedTo = sortAlphabetical([...state.to]);
  const sortedCc = sortAlphabetical([...state.cc]);
  const sortedBcc = sortAlphabetical([...state.bcc]);

  const action = {
    type: 'sort',
    input: input,
    output: {
      to: sortedTo,
      cc: sortedCc,
      bcc: sortedBcc
    },
    processed: input.to.length + input.cc.length + input.bcc.length
  };

  return {
    ...state,
    to: sortedTo,
    cc: sortedCc,
    bcc: sortedBcc,
    actions: [...state.actions, action]
  };
}

// ============================================================================
// DEDUPE FUNCTIONS
// ============================================================================

function removeDuplicates(recipients) {
  const seen = new Set();
  const removed = [];

  const unique = recipients.filter(recipient => {
    const email = extractEmail(recipient).toLowerCase();
    if (seen.has(email)) {
      removed.push(recipient);
      return false;
    }
    seen.add(email);
    return true;
  });

  return { unique, removed };
}

function dedupeArrays(to, cc, bcc) {
  const { unique: uniqueTo, removed: removedFromTo } = removeDuplicates(to);
  const { unique: uniqueCc, removed: removedFromCc } = removeDuplicates(cc);
  const { unique: uniqueBcc, removed: removedFromBcc } = removeDuplicates(bcc);

  const allEmails = new Set();
  const finalTo = [];
  const finalCc = [];
  const finalBcc = [];
  const crossArrayRemoved = [];

  uniqueTo.forEach(recipient => {
    const email = extractEmail(recipient).toLowerCase();
    if (!allEmails.has(email)) {
      allEmails.add(email);
      finalTo.push(recipient);
    } else {
      crossArrayRemoved.push(recipient);
    }
  });

  uniqueCc.forEach(recipient => {
    const email = extractEmail(recipient).toLowerCase();
    if (!allEmails.has(email)) {
      allEmails.add(email);
      finalCc.push(recipient);
    } else {
      crossArrayRemoved.push(recipient);
    }
  });

  uniqueBcc.forEach(recipient => {
    const email = extractEmail(recipient).toLowerCase();
    if (!allEmails.has(email)) {
      allEmails.add(email);
      finalBcc.push(recipient);
    } else {
      crossArrayRemoved.push(recipient);
    }
  });

  return {
    to: finalTo,
    cc: finalCc,
    bcc: finalBcc,
    removed: [
      ...removedFromTo,
      ...removedFromCc,
      ...removedFromBcc,
      ...crossArrayRemoved
    ]
  };
}

function dedupeStep(state) {
  const input = {
    to: [...state.to],
    cc: [...state.cc],
    bcc: [...state.bcc]
  };

  const result = dedupeArrays(state.to, state.cc, state.bcc);

  const action = {
    type: 'dedupe',
    input: input,
    removed: result.removed,
    output: {
      to: result.to,
      cc: result.cc,
      bcc: result.bcc
    },
    processed: input.to.length + input.cc.length + input.bcc.length,
    duplicatesFound: result.removed.length
  };

  return {
    ...state,
    to: result.to,
    cc: result.cc,
    bcc: result.bcc,
    actions: [...state.actions, action]
  };
}

// ============================================================================
// VALIDATION FUNCTIONS
// ============================================================================

const COMMON_TYPOS = {
  'gmial.com': 'gmail.com',
  'gmai.com': 'gmail.com',
  'gmal.com': 'gmail.com',
  'yahooo.com': 'yahoo.com',
  'yaho.com': 'yahoo.com',
  'hotmial.com': 'hotmail.com',
  'hotmial.co': 'hotmail.com',
  'outlok.com': 'outlook.com',
  'outloo.com': 'outlook.com',
  'microsft.com': 'microsoft.com',
  'mircosoft.com': 'microsoft.com'
};

function validateEmailFormat(email) {
  if (!email || typeof email !== 'string') {
    return { isValid: false, message: 'Empty or invalid email address' };
  }

  email = email.trim();

  if (email.length <= 3 && !email.includes('@')) {
    return { isValid: false, message: 'Email address is too short or malformed' };
  }

  if (email.endsWith(':') && !email.includes('@')) {
    return { isValid: false, message: 'Email appears to be incomplete or malformed' };
  }

  if (email.length === 0) {
    return { isValid: false, message: 'Empty email address' };
  }

  if (!email.includes('@')) {
    return { isValid: false, message: 'Email must contain @ symbol' };
  }

  if (email.split('@').length !== 2) {
    return { isValid: false, message: 'Email must contain exactly one @ symbol' };
  }

  const [localPart, domainPart] = email.split('@');

  if (!localPart || localPart.length === 0) {
    return { isValid: false, message: 'Email must have content before @ symbol' };
  }

  if (localPart.length > 64) {
    return { isValid: false, message: 'Local part of email is too long' };
  }

  if (!domainPart || domainPart.length === 0) {
    return { isValid: false, message: 'Email must have content after @ symbol' };
  }

  if (domainPart.length > 253) {
    return { isValid: false, message: 'Domain part of email is too long' };
  }

  if (!domainPart.includes('.')) {
    return { isValid: false, message: 'Domain must contain at least one dot' };
  }

  const emailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;

  if (!emailRegex.test(email)) {
    return { isValid: false, message: 'Invalid email format' };
  }

  if (email.includes('..')) {
    return { isValid: false, message: 'Email contains consecutive dots' };
  }

  if (email.startsWith('.') || email.endsWith('.')) {
    return { isValid: false, message: 'Email starts or ends with a dot' };
  }

  if (email.includes('@.') || email.includes('.@')) {
    return { isValid: false, message: 'Invalid dot placement near @ symbol' };
  }

  if (email.startsWith('-') || email.endsWith('-')) {
    return { isValid: false, message: 'Email starts or ends with hyphen' };
  }

  const domainParts = domainPart.split('.');
  const tld = domainParts[domainParts.length - 1];
  if (tld.length < 2) {
    return { isValid: false, message: 'Top-level domain must be at least 2 characters' };
  }

  return { isValid: true };
}

function calculateSimilarity(str1, str2) {
  const longer = str1.length > str2.length ? str1 : str2;
  const shorter = str1.length > str2.length ? str2 : str1;

  if (longer.length === 0) {
    return 1.0;
  }

  const editDistance = getEditDistance(longer, shorter);
  return (longer.length - editDistance) / longer.length;
}

function getEditDistance(str1, str2) {
  const matrix = [];

  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }

  return matrix[str2.length][str1.length];
}

function checkForTypos(email) {
  const domain = email.split('@')[1];
  if (!domain) {
    return { hasTypo: false };
  }

  const lowerDomain = domain.toLowerCase();

  if (COMMON_TYPOS[lowerDomain]) {
    const correctedEmail = email.replace(
      new RegExp(domain.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i'),
      COMMON_TYPOS[lowerDomain]
    );

    return {
      hasTypo: true,
      suggestion: correctedEmail,
      message: `Did you mean "${correctedEmail}"?`
    };
  }

  const commonDomains = ['gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'microsoft.com'];

  for (const commonDomain of commonDomains) {
    if (Math.abs(lowerDomain.length - commonDomain.length) <= 2) {
      const similarity = calculateSimilarity(lowerDomain, commonDomain);
      if (similarity > 0.8) {
        const correctedEmail = email.replace(
          new RegExp(domain.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i'),
          commonDomain
        );

        return {
          hasTypo: true,
          suggestion: correctedEmail,
          message: `Did you mean "${correctedEmail}"?`
        };
      }
    }
  }

  return { hasTypo: false };
}

function validateRecipient(recipient) {
  const email = extractEmail(recipient);
  const results = {
    address: recipient,
    email: email,
    isValid: true,
    status: 'valid',
    warnings: [],
    suggestions: []
  };

  const formatCheck = validateEmailFormat(email);
  if (!formatCheck.isValid) {
    results.isValid = false;
    results.status = 'error';
    results.warnings.push(formatCheck.message);
    return results;
  }

  const typoCheck = checkForTypos(email);
  if (typoCheck.hasTypo) {
    results.status = 'warning';
    results.warnings.push(typoCheck.message);
    results.suggestions.push(typoCheck.suggestion);
  }

  return results;
}

function validateAndFilterRecipients(recipients) {
  const results = [];
  const validRecipients = [];

  for (const recipient of recipients) {
    try {
      const result = validateRecipient(recipient);
      results.push(result);

      if (result.status === 'valid' || result.status === 'warning') {
        validRecipients.push(recipient);
      }
    } catch (error) {
      const errorResult = {
        address: recipient,
        email: extractEmail(recipient),
        status: 'error',
        warnings: ['Validation failed: ' + error.message]
      };
      results.push(errorResult);
    }
  }

  return { results, validRecipients };
}

function validateStep(state) {
  const input = {
    to: [...state.to],
    cc: [...state.cc],
    bcc: [...state.bcc]
  };

  const processedTo = validateAndFilterRecipients(state.to);
  const processedCc = validateAndFilterRecipients(state.cc);
  const processedBcc = validateAndFilterRecipients(state.bcc);

  const allValidationResults = [
    ...processedTo.results,
    ...processedCc.results,
    ...processedBcc.results
  ];

  const errors = allValidationResults.filter(r => r.status === 'error');
  const warnings = allValidationResults.filter(r => r.status === 'warning');
  const valid = allValidationResults.filter(r => r.status === 'valid');

  const action = {
    type: 'validate',
    input: input,
    output: {
      to: processedTo.validRecipients,
      cc: processedCc.validRecipients,
      bcc: processedBcc.validRecipients
    },
    validationResults: allValidationResults,
    errors: errors,
    warnings: warnings,
    processed: state.to.length + state.cc.length + state.bcc.length,
    validCount: valid.length,
    warningCount: warnings.length,
    errorCount: errors.length,
    removedInvalid: errors.length
  };

  return {
    ...state,
    to: processedTo.validRecipients,
    cc: processedCc.validRecipients,
    bcc: processedBcc.validRecipients,
    actions: [...state.actions, action]
  };
}

// ============================================================================
// PRIORITIZE INTERNAL FUNCTIONS
// ============================================================================

function getDomainIndex(email, internalDomains) {
  if (!email || !email.includes('@')) {
    return -1;
  }

  const domain = email.split('@')[1].toLowerCase();

  for (let i = 0; i < internalDomains.length; i++) {
    const internalDomain = internalDomains[i].toLowerCase();
    if (domain === internalDomain || domain.endsWith('.' + internalDomain)) {
      return i;
    }
  }

  return -1;
}

function prioritizeInternal(recipients, internalDomains, sortAlphabetically = false) {
  if (!recipients || recipients.length === 0) {
    return recipients;
  }

  if (!internalDomains || internalDomains.length === 0) {
    return recipients;
  }

  const internal = [];
  const external = [];

  recipients.forEach(recipient => {
    const email = extractEmail(recipient);
    const domainIndex = getDomainIndex(email, internalDomains);

    if (domainIndex >= 0) {
      internal.push({ recipient, domainIndex, email });
    } else {
      external.push({ recipient, email });
    }
  });

  internal.sort((a, b) => {
    if (a.domainIndex !== b.domainIndex) {
      return a.domainIndex - b.domainIndex;
    }

    if (sortAlphabetically) {
      const nameA = extractDisplayName(a.recipient).toLowerCase();
      const nameB = extractDisplayName(b.recipient).toLowerCase();

      if (!nameA && !nameB) {
        return a.email.toLowerCase().localeCompare(b.email.toLowerCase());
      }

      return nameA.localeCompare(nameB);
    }

    return 0;
  });

  if (sortAlphabetically) {
    external.sort((a, b) => {
      const nameA = extractDisplayName(a.recipient).toLowerCase();
      const nameB = extractDisplayName(b.recipient).toLowerCase();

      if (!nameA && !nameB) {
        return a.email.toLowerCase().localeCompare(b.email.toLowerCase());
      }

      return nameA.localeCompare(nameB);
    });
  }

  return [
    ...internal.map(item => item.recipient),
    ...external.map(item => item.recipient)
  ];
}

function prioritizeInternalStep(state) {
  const input = {
    to: [...state.to],
    cc: [...state.cc],
    bcc: [...state.bcc]
  };

  const internalDomains = state.internalDomains || [];
  const sortAlphabetically = state.enabledSteps?.includes('sort') || false;

  const prioritizedTo = prioritizeInternal([...state.to], internalDomains, sortAlphabetically);
  const prioritizedCc = prioritizeInternal([...state.cc], internalDomains, sortAlphabetically);
  const prioritizedBcc = prioritizeInternal([...state.bcc], internalDomains, sortAlphabetically);

  const action = {
    type: 'prioritizeInternal',
    input: input,
    output: {
      to: prioritizedTo,
      cc: prioritizedCc,
      bcc: prioritizedBcc
    },
    processed: input.to.length + input.cc.length + input.bcc.length
  };

  return {
    ...state,
    to: prioritizedTo,
    cc: prioritizedCc,
    bcc: prioritizedBcc,
    actions: [...state.actions, action]
  };
}

// ============================================================================
// REMOVE EXTERNAL FUNCTIONS
// ============================================================================

function isInternalDomain(email, internalDomains) {
  if (!email || !email.includes('@')) {
    return false;
  }

  const domain = email.split('@')[1].toLowerCase();

  for (const internalDomain of internalDomains) {
    const internal = internalDomain.toLowerCase();
    if (domain === internal || domain.endsWith('.' + internal)) {
      return true;
    }
  }

  return false;
}

function removeExternal(recipients, internalDomains) {
  if (!recipients || recipients.length === 0) {
    return { filtered: [], removed: [] };
  }

  if (!internalDomains || internalDomains.length === 0) {
    return { filtered: [...recipients], removed: [] };
  }

  const filtered = [];
  const removed = [];

  recipients.forEach(recipient => {
    const email = extractEmail(recipient);

    if (isInternalDomain(email, internalDomains)) {
      filtered.push(recipient);
    } else {
      removed.push(recipient);
    }
  });

  return { filtered, removed };
}

function removeExternalStep(state) {
  const input = {
    to: [...state.to],
    cc: [...state.cc],
    bcc: [...state.bcc]
  };

  const internalDomains = state.internalDomains || [];

  const toResult = removeExternal([...state.to], internalDomains);
  const ccResult = removeExternal([...state.cc], internalDomains);
  const bccResult = removeExternal([...state.bcc], internalDomains);

  const totalRemoved = toResult.removed.length + ccResult.removed.length + bccResult.removed.length;

  const action = {
    type: 'removeExternal',
    input: input,
    output: {
      to: toResult.filtered,
      cc: ccResult.filtered,
      bcc: bccResult.filtered
    },
    removed: {
      to: toResult.removed,
      cc: ccResult.removed,
      bcc: bccResult.removed,
      total: totalRemoved
    },
    processed: input.to.length + input.cc.length + input.bcc.length
  };

  return {
    ...state,
    to: toResult.filtered,
    cc: ccResult.filtered,
    bcc: bccResult.filtered,
    actions: [...state.actions, action]
  };
}

// ============================================================================
// FLAG EXTERNAL FUNCTIONS
// ============================================================================

function isExternalEmail(email, orgDomain) {
  if (!email || !orgDomain) {
    return false;
  }

  const emailDomain = email.split('@')[1];
  if (!emailDomain) {
    return true;
  }

  const normalizedEmailDomain = emailDomain.toLowerCase();
  const normalizedOrgDomain = orgDomain.toLowerCase();

  if (normalizedEmailDomain === normalizedOrgDomain) {
    return false;
  }

  if (normalizedEmailDomain.endsWith('.' + normalizedOrgDomain)) {
    return false;
  }

  return true;
}

function categorizeRecipients(recipients, orgDomain) {
  const internal = [];
  const external = [];

  recipients.forEach(recipient => {
    const email = extractEmail(recipient);
    if (isExternalEmail(email, orgDomain)) {
      external.push({
        recipient,
        email,
        domain: email.split('@')[1] || 'unknown',
        isExternal: true
      });
    } else {
      internal.push({
        recipient,
        email,
        domain: email.split('@')[1] || 'unknown',
        isExternal: false
      });
    }
  });

  return { internal, external };
}

function flagExternalStep(state) {
  const { orgDomain } = state;

  if (!orgDomain) {
    const action = {
      type: 'flagExt',
      input: {
        to: [...state.to],
        cc: [...state.cc],
        bcc: [...state.bcc]
      },
      output: {
        to: state.to,
        cc: state.cc,
        bcc: state.bcc
      },
      flagged: [],
      processed: 0,
      skipped: true,
      message: 'No organization domain provided'
    };

    return {
      ...state,
      actions: [...state.actions, action]
    };
  }

  const input = {
    to: [...state.to],
    cc: [...state.cc],
    bcc: [...state.bcc]
  };

  const toCategorized = categorizeRecipients(state.to, orgDomain);
  const ccCategorized = categorizeRecipients(state.cc, orgDomain);
  const bccCategorized = categorizeRecipients(state.bcc, orgDomain);

  const allExternal = [
    ...toCategorized.external.map(r => ({ ...r, field: 'to' })),
    ...ccCategorized.external.map(r => ({ ...r, field: 'cc' })),
    ...bccCategorized.external.map(r => ({ ...r, field: 'bcc' }))
  ];

  const externalDomains = allExternal.reduce((acc, recipient) => {
    const domain = recipient.domain;
    if (!acc[domain]) {
      acc[domain] = [];
    }
    acc[domain].push(recipient);
    return acc;
  }, {});

  const action = {
    type: 'flagExt',
    input: input,
    output: {
      to: state.to,
      cc: state.cc,
      bcc: state.bcc
    },
    flagged: allExternal.map(r => r.recipient),
    externalByDomain: externalDomains,
    summary: {
      totalRecipients: state.to.length + state.cc.length + state.bcc.length,
      externalCount: allExternal.length,
      internalCount: (state.to.length + state.cc.length + state.bcc.length) - allExternal.length,
      uniqueExternalDomains: Object.keys(externalDomains).length
    },
    orgDomain: orgDomain,
    processed: state.to.length + state.cc.length + state.bcc.length
  };

  return {
    ...state,
    actions: [...state.actions, action]
  };
}

// ============================================================================
// MAIN ORCHESTRATOR
// ============================================================================

function processRecipients(payload) {
  const stepMap = {
    'sort': sortStep,
    'dedupe': dedupeStep,
    'validate': validateStep,
    'prioritizeInternal': prioritizeInternalStep,
    'removeExternal': removeExternalStep,
    'flagExt': flagExternalStep
  };

  let state = {
    to: payload.to || [],
    cc: payload.cc || [],
    bcc: payload.bcc || [],
    enabledSteps: payload.userSettings?.enabledSteps || [],
    internalDomains: payload.userSettings?.internalDomains || [],
    orgDomain: payload.userSettings?.orgDomain || '',
    actions: []
  };

  for (const stepName of state.enabledSteps) {
    const stepFn = stepMap[stepName];
    if (stepFn) {
      try {
        state = stepFn(state);
      } catch (error) {
        throw new Error(`Step '${stepName}' failed: ${error.message}`);
      }
    }
  }

  return {
    success: true,
    result: {
      to: state.to,
      cc: state.cc,
      bcc: state.bcc
    },
    actions: state.actions,
    summary: {
      totalProcessed: (payload.to?.length || 0) + (payload.cc?.length || 0) + (payload.bcc?.length || 0),
      totalRemaining: state.to.length + state.cc.length + state.bcc.length,
      stepsExecuted: state.actions.length
    }
  };
}

// Export for use in taskpane.js
if (typeof window !== 'undefined') {
  window.ClearSendProcessors = {
    processRecipients,
    extractEmail,
    extractDisplayName
  };
}
