function validateCohortSheetHeaders(internshipId, cohortName) {
  Logger.log(`Starting header validation for Internship ID: ${internshipId}, Cohort: ${cohortName}`);
  let validationResult = {
    success: true,
    message: 'Sheet headers validated successfully.',
    details: {
      missingCritical: [],
      missingRecommended: [],
      foundWeeks: false
    }
  };

  try {
    if (!internshipId || !cohortName) {
      throw new Error("Internship ID and Cohort Name are required for header validation.");
    }

    const ss = SpreadsheetApp.openById(internshipId);
    const cohortSheet = ss.getSheetByName(cohortName);

    if (!cohortSheet) {
      throw new Error(`Cohort sheet "${cohortName}" not found in internship spreadsheet (ID: ${internshipId}).`);
    }

    const headers = cohortSheet.getRange(1, 1, 1, cohortSheet.getLastColumn()).getValues()[0];
    const headerSet = new Set(headers.map(h => typeof h === 'string' ? h.trim() : '')); 

    const criticalHeaders = ['Full Name', 'Email Address'];
    criticalHeaders.forEach(header => {
      if (!headerSet.has(header)) {
        validationResult.success = false;
        validationResult.details.missingCritical.push(header);
      }
    });

    const recommendedHeaders = [
      'Submitted Feedback Form',
      'Submitted Reflective Video',
      'Submitted Peer Evaluation Form',
      'Submitted Self Evaluation Form'
    ];
    recommendedHeaders.forEach(header => {
      if (!headerSet.has(header)) {
        validationResult.details.missingRecommended.push(header);
      }
    });

    const hasWeekColumn = headers.some(header => typeof header === 'string' && header.toLowerCase().includes('week'));
    if (hasWeekColumn) {
      validationResult.details.foundWeeks = true;
    } else {
      validationResult.details.missingRecommended.push("At least one 'Week X' column");
    }

    const hasStatusColumn = headers.some(header => typeof header === 'string' && header.toLowerCase().includes('status'));
    if (!hasStatusColumn) {
      validationResult.details.missingRecommended.push("Status (for skipping completed interns)");
    }

    if (!validationResult.success) {
      validationResult.message = `Validation FAILED: Missing critical column(s): ${validationResult.details.missingCritical.join(', ')}. Please add them to the sheet.`;
      Logger.log(validationResult.message);
    } else {
      if (validationResult.details.missingRecommended.length > 0) {
        validationResult.message += ` Optional/Recommended columns missing: ${validationResult.details.missingRecommended.join(', ')}.`;
        Logger.log(validationResult.message);
      }
      if (!validationResult.details.foundWeeks) {
         validationResult.message += ` No 'Week X' columns found. No weekly deliverables will be checked.`;
         Logger.log(validationResult.message);
      }
      Logger.log('Sheet headers validated successfully with no critical issues.');
    }

  } catch (e) {
    validationResult.success = false;
    validationResult.message = `Error during sheet header validation: ${e.message}. Please check the Internship ID and Cohort Name.`;
    Logger.log(validationResult.message);
  }

  return validationResult;
}
