// Configuration
const AFFINDA_API_KEY = [your api]
const SHEET_NAME = 'Sheet1';

// Main function to process resumes
function processResumes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found. Please check the sheet name.');
  }
  
  const lastRow = sheet.getLastRow();
  
  // Set headers if not present
  const headers = sheet.getRange('J1:L1').getValues()[0];
  if (!headers[0]) sheet.getRange('J1').setValue('Upload your updated CV');
  if (!headers[1]) sheet.getRange('K1').setValue('JD');
  if (!headers[2]) sheet.getRange('L1').setValue('ATS Score');
  
  // Start from row 2 to skip header
  for (let row = 2; row <= lastRow; row++) {
    const resumeLink = sheet.getRange(`J${row}`).getValue();
    const jobDescription = sheet.getRange(`K${row}`).getValue();
    
    if (resumeLink && jobDescription) {
      try {
        // First try to get the file from Google Drive
        const fileId = getDriveFileIdFromUrl(resumeLink);
        if (!fileId) {
          sheet.getRange(`L${row}`).setValue('Invalid Drive URL');
          continue;
        }
        
        const file = DriveApp.getFileById(fileId);
        if (!file) {
          sheet.getRange(`L${row}`).setValue('File not found in Drive');
          continue;
        }
        
        const score = calculateATSScore(file.getBlob(), jobDescription);
        sheet.getRange(`L${row}`).setValue(score);
      } catch (error) {
        Logger.log(`Error processing row ${row}: ${error.message}`);
        sheet.getRange(`L${row}`).setValue('Error: ' + error.message);
      }
    }
  }
}

// Function to parse resume using Affinda API
function parseResumeWithAffinda(resumeBlob) {
  const endpoint = 'https://api.affinda.com/v2/resumes';
  
  const options = {
    'method': 'POST',
    'headers': {
      'Authorization': `Bearer ${AFFINDA_API_KEY}`,
      'Accept': 'application/json'
    },
    'payload': {
      'file': resumeBlob,
      'wait': 'true'
    },
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      Logger.log('API Response: ' + responseText);
      throw new Error(responseText);
    }
    
    return JSON.parse(responseText);
  } catch (error) {
    Logger.log('Error parsing resume: ' + error.message);
    throw new Error('Resume parsing failed: ' + error.message);
  }
}

// Helper function to get Drive file ID from URL
function getDriveFileIdFromUrl(url) {
  // Handle both sharing URL formats
  const patterns = [
    /\/d\/([-\w]{25,})/,
    /id=([-\w]{25,})/,
    /\/file\/d\/([-\w]{25,})/
  ];
  
  for (let pattern of patterns) {
    const match = url.match(pattern);
    if (match) return match[1];
  }
  return null;
}

// Function to calculate matching score between resume and job description
function calculateATSScore(resumeBlob, jobDescription) {
  // Parse resume first
  const resumeData = parseResumeWithAffinda(resumeBlob);
  
  try {
    // Extract skills from resume data
    const resumeSkills = resumeData.data.skills.map(skill => skill.name.toLowerCase()) || [];
    
    // Extract requirements from job description
    const requirements = extractRequirements(jobDescription);
    
    // Calculate matching score
    let matchingPoints = 0;
    let totalPoints = requirements.length;
    
    requirements.forEach(requirement => {
      const req = requirement.toLowerCase();
      if (resumeSkills.some(skill => skill.includes(req) || req.includes(skill))) {
        matchingPoints++;
      }
    });
    
    // Convert to percentage
    const score = (matchingPoints / totalPoints) * 100;
    return Math.round(score) + '%';
  } catch (error) {
    Logger.log('Error calculating score: ' + error.message);
    throw new Error('Score calculation failed: ' + error.message);
  }
}

// Helper function to extract requirements from job description
function extractRequirements(jobDescription) {
  // Split the job description into words and clean them
  const words = jobDescription.split(/[\s,;]+/);
  const cleanedWords = words.map(word => 
    word.replace(/[^\w\s+#]/g, '').trim() // Keep +, # and letters/numbers
  ).filter(word => 
    word.length > 2 && // Keep words longer than 2 characters
    !/^\d+$/.test(word) // Remove pure numbers
  );
  
  // Remove duplicates
  return [...new Set(cleanedWords)];
}

// Add menu item to sheet
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ATS Processing')
    .addItem('Process Resumes', 'processResumes')
    .addToUi();
} 
