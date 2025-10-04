/**
 * Canvas Assignment Tracker — Streamlined Script
 * - Canvas API comprehensive import (assignments, grades, weights, submission status)
 * - Automatic data cleaning (default times, points extraction, days left calculation)
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Canvas Tracker')
    .addItem('Import Everything from Canvas', 'uiImportFromCanvasApi')
    .addSeparator()
    .addItem('Refresh Dashboard', 'uiRefreshDashboard')
    .addSeparator()
    .addItem('Set Canvas Domain', 'uiSetCanvasDomain')
    .addItem('Set Canvas API Token', 'uiSetCanvasToken')
    .addToUi();
}

// ==== CONFIG ====
const RAW_DATA_SHEET = 'Raw Data';
const DASHBOARD_SHEET = 'Dashboard';


// ==== HELPERS ====
function getRawDataSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(RAW_DATA_SHEET);
  if (!sh) sh = ss.insertSheet(RAW_DATA_SHEET);
  
  const headers = [
    // Core Assignment Fields
    'id', 'name', 'description', 'course_id', 'assignment_group_id', 'position',
    // Dates & Times
    'due_at', 'unlock_at', 'lock_at', 'created_at', 'updated_at',
    // Points & Grading
    'points_possible', 'grading_type', 'grading_standard_id', 'omit_from_final_grade',
    // Submission Settings
    'submission_types', 'allowed_extensions', 'allowed_attempts', 'annotatable_attachment_id',
    // Status & Visibility
    'published', 'muted', 'anonymous_submissions', 'anonymous_grading', 'anonymous_instructor_annotations',
    'hide_in_gradebook', 'post_to_sis', 'integration_id', 'integration_data',
    // Peer Reviews & Moderation
    'peer_reviews', 'automatic_peer_reviews', 'peer_review_count', 'peer_reviews_assign_at',
    'intra_group_peer_reviews', 'moderated_grading', 'grader_count', 'final_grader_id',
    'grader_comments_visible_to_graders', 'graders_anonymous_to_graders', 'grader_names_visible_to_final_grader',
    // External Tools & Plagiarism
    'turnitin_enabled', 'vericite_enabled', 'turnitin_settings', 'external_tool_tag_attributes',
    // Rubrics
    'use_rubric_for_grading', 'rubric_settings', 'rubric', 'free_form_criterion_comments',
    // URLs & Links
    'html_url', 'submissions_download_url', 'quiz_id', 'discussion_topic',
    // Assignment Group Info (from separate API call)
    'assignment_group_name', 'group_weight',
    // Submission Status (from submissions API)
    'submission_workflow_state', 'submission_score', 'submission_grade', 'submission_submitted_at',
    'submission_graded_at', 'submission_late', 'submission_missing', 'submission_excused',
    // Calculated Fields
    'days_until_due', 'is_overdue'
  ];
  
  // Handle header setup - if sheet is empty or has wrong number of columns, reset headers
  const currentCols = sh.getLastColumn();
  const needsCols = headers.length;
  
  if (sh.getLastRow() === 0) {
    // Empty sheet, just add headers
    sh.appendRow(headers);
  } else if (currentCols !== needsCols) {
    // Existing sheet with wrong column count - update headers
    if (currentCols < needsCols) {
      sh.insertColumns(currentCols + 1, needsCols - currentCols);
    }
    sh.getRange(1, 1, 1, needsCols).setValues([headers]);
  }
  
  return sh;
}

function getDashboardSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DASHBOARD_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DASHBOARD_SHEET);
    setupDashboard_(sh);
  }
  return sh;
}

function setupDashboard_(sh) {
  // Clear existing content
  sh.clear();
  
  // Set up dashboard layout
  const headers = [
    // Row 1: Title
    ['CANVAS ASSIGNMENT TRACKER DASHBOARD', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    
    // Row 3: Course Summary Headers
    ['COURSE SUMMARY', '', '', 'UPCOMING DEADLINES', '', '', 'OVERDUE ASSIGNMENTS', ''],
    ['Course', 'Current Grade', 'Assignments Due', 'Assignment', 'Course', 'Due Date', 'Assignment', 'Days Overdue'],
    
    // Rows 5-14: Data rows (will be populated by formulas)
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    
    // Row 15: Grade Analytics Headers
    ['', '', '', '', '', '', '', ''],
    ['GRADE ANALYTICS', '', '', 'ASSIGNMENT TYPE PERFORMANCE', '', '', 'WORKLOAD ANALYSIS', ''],
    ['Total Points Earned', 'Total Points Possible', 'Overall GPA', 'Type', 'Avg Score', 'Count', 'This Week', 'Next Week'],
    ['', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '']
  ];
  
  // Set the data
  sh.getRange(1, 1, headers.length, 8).setValues(headers);
  
  // Format the dashboard
  formatDashboard_(sh);
}

function formatDashboard_(sh) {
  // Title formatting
  sh.getRange('A1:H1').merge().setHorizontalAlignment('center')
    .setFontSize(16).setFontWeight('bold')
    .setBackground('#4285f4').setFontColor('white');
  
  // Section headers
  sh.getRange('A3').setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange('D3').setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange('G3').setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange('A16').setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange('D16').setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange('G16').setFontWeight('bold').setBackground('#e8f0fe');
  
  // Column headers
  sh.getRange('A4:H4').setFontWeight('bold').setBackground('#f8f9fa');
  sh.getRange('A17:H17').setFontWeight('bold').setBackground('#f8f9fa');
  
  // Auto-resize columns
  sh.autoResizeColumns(1, 8);
  
  // Add borders
  sh.getRange('A1:H19').setBorder(true, true, true, true, true, true);
}

function updateDashboardFormulas_(dashboardSheet) {
  const rawSheetName = RAW_DATA_SHEET;
  
  // Use simpler, more compatible formulas
  try {
    // Course Summary - Manually list some course IDs from your data
    const courseIds = ['43899', '51304', '53553', '51269', '53554', '26518', '53480', '55444'];
    
    // Set course IDs manually for now
    for (let i = 0; i < courseIds.length; i++) {
      dashboardSheet.getRange(5 + i, 1).setValue(courseIds[i]);
      
      // Current Grade - Average submission scores for this course
      dashboardSheet.getRange(5 + i, 2).setFormula(`=IFERROR(AVERAGEIF('${rawSheetName}'!D:D,"${courseIds[i]}",'${rawSheetName}'!BC:BC),"No grades")`);
      
      // Assignments Due - Count upcoming assignments for this course
      dashboardSheet.getRange(5 + i, 3).setFormula(`=COUNTIFS('${rawSheetName}'!D:D,"${courseIds[i]}",'${rawSheetName}'!BH:BH,">0")`);
    }
    
    // Grade Analytics - Total Points Earned
    dashboardSheet.getRange('A18').setFormula(`=SUMPRODUCT(('${rawSheetName}'!BC:BC>0)*('${rawSheetName}'!BC:BC))`);
    
    // Total Points Possible
    dashboardSheet.getRange('B18').setFormula(`=SUMPRODUCT(('${rawSheetName}'!L:L>0)*('${rawSheetName}'!L:L))`);
    
    // Overall GPA (percentage)
    dashboardSheet.getRange('C18').setFormula(`=IF(B18>0,A18/B18*100,"No data")`);
    
    // Workload Analysis - This week (0-7 days)
    dashboardSheet.getRange('G18').setFormula(`=COUNTIFS('${rawSheetName}'!BH:BH,">=0",'${rawSheetName}'!BH:BH,"<=7")`);
    
    // Next week (8-14 days)  
    dashboardSheet.getRange('H18').setFormula(`=COUNTIFS('${rawSheetName}'!BH:BH,">7",'${rawSheetName}'!BH:BH,"<=14")`);
    
    // Add some upcoming deadlines manually
    dashboardSheet.getRange('D5').setFormula(`=INDEX(FILTER('${rawSheetName}'!B:B,'${rawSheetName}'!BH:BH>0,'${rawSheetName}'!BH:BH<=7),1)`);
    dashboardSheet.getRange('E5').setFormula(`=INDEX(FILTER('${rawSheetName}'!D:D,'${rawSheetName}'!BH:BH>0,'${rawSheetName}'!BH:BH<=7),1)`);
    dashboardSheet.getRange('F5').setFormula(`=INDEX(FILTER('${rawSheetName}'!BH:BH,'${rawSheetName}'!BH:BH>0,'${rawSheetName}'!BH:BH<=7),1)`);
    
    // Add some overdue assignments
    dashboardSheet.getRange('G5').setFormula(`=INDEX(FILTER('${rawSheetName}'!B:B,'${rawSheetName}'!BH:BH<0),1)`);
    dashboardSheet.getRange('H5').setFormula(`=INDEX(FILTER('${rawSheetName}'!BH:BH,'${rawSheetName}'!BH:BH<0),1)`);
    
  } catch (error) {
    console.log('Error applying dashboard formulas:', error);
    
    // Ultimate fallback - just show some basic stats
    dashboardSheet.getRange('A18').setFormula(`=SUM('${rawSheetName}'!BC:BC)`);
    dashboardSheet.getRange('B18').setFormula(`=SUM('${rawSheetName}'!L:L)`);
    dashboardSheet.getRange('G18').setFormula(`=COUNTIF('${rawSheetName}'!BH:BH,"<=7")`);
    dashboardSheet.getRange('H18').setFormula(`=COUNTIF('${rawSheetName}'!BH:BH,">7")`);
  }
}

function uiRefreshDashboard() {
  const dashboardSheet = getDashboardSheet();
  updateDashboardFormulas_(dashboardSheet);
  SpreadsheetApp.getUi().alert('Dashboard refreshed with latest data!');
}
function getBodyRange_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1) return null;
  return sh.getRange(2, 1, lastRow - 1, lastCol);
}
function writeRowsInBatches_(sh, startRow, values, batchSize) {
  if (!values.length) return;
  
  const bs = batchSize || 300;
  const dataColCount = values[0].length;
  const sheetColCount = sh.getLastColumn();
  
  // If data has more columns than sheet, expand the sheet
  if (dataColCount > sheetColCount) {
    sh.insertColumns(sheetColCount + 1, dataColCount - sheetColCount);
  }
  
  for (let i = 0; i < values.length; i += bs) {
    const slice = values.slice(i, i + bs);
    sh.getRange(startRow + i, 1, slice.length, dataColCount).setValues(slice);
    SpreadsheetApp.flush();
    Utilities.sleep(30);
  }
}
function stripHtml_(html) {
  if (!html) return '';
  return html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}
function guessTypeFromSummary(s) {
  s = (s || '').toLowerCase();
  if (s.includes('quiz')) return 'Quiz';
  if (s.includes('exam') || s.includes('midterm') || s.includes('final')) return 'Exam';
  if (s.includes('lab')) return 'Lab';
  if (s.includes('project')) return 'Project';
  if (s.includes('hw') || s.includes('homework') || s.includes('assignment')) return 'Homework';
  if (s.includes('reading')) return 'Reading';
  return 'Other';
}

// ==== CLEANING HELPERS ====
function autofillAndClean_() {
  const sh = getSheet();
  const body = getBodyRange_(sh);
  if (!body) return;

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx = (n) => header.indexOf(n);

  const dueDateIdx  = idx('Due Date');
  const dueTimeIdx  = idx('Due Time');
  const typeIdx     = idx('Type');
  const categoryIdx = idx('Category');
  const pointsIdx   = idx('Points');
  const notesIdx    = idx('Notes');
  const daysLeftIdx = idx('Days Left');

  const values = body.getValues();
  const pointsRegex = /(\d+(?:\.\d+)?)\s*(?:pts?|points?)/i;

  const bs = 300;
  for (let off = 0; off < values.length; off += bs) {
    const slice = values.slice(off, off + bs);
    for (let r = 0; r < slice.length; r++) {
      // Fix missing due times
      const t = (slice[r][dueTimeIdx] || '').toString();
      if (!t || t === '0:00' || t === '00:00') slice[r][dueTimeIdx] = '23:59';

      // Copy type to category if category is empty
      const ty = (slice[r][typeIdx] || '').toString();
      if (ty && !slice[r][categoryIdx]) slice[r][categoryIdx] = ty;

      // Extract points from notes if points field is empty
      const notes = (slice[r][notesIdx] || '').toString();
      const pm = notes.match(pointsRegex);
      if (pm && !slice[r][pointsIdx]) slice[r][pointsIdx] = pm[1];

      // Add days left formula
      if (!slice[r][daysLeftIdx]) {
        const sheetRow = 2 + off + r;
        const dueCell = sh.getRange(sheetRow, dueDateIdx + 1).getA1Notation();
        slice[r][daysLeftIdx] = `=IF(${dueCell}="","", ${dueCell} - TODAY())`;
      }
    }
    sh.getRange(2 + off, 1, slice.length, header.length).setValues(slice);
    SpreadsheetApp.flush();
    Utilities.sleep(20);
  }
}

// ==== CANVAS API (points, links) ====
// store per-user domain/token in UserProperties
function uiSetCanvasDomain() {
  const ui = SpreadsheetApp.getUi();
  const def = PropertiesService.getUserProperties().getProperty('CANVAS_DOMAIN') || 'canvas.gmu.edu';
  const resp = ui.prompt('Canvas domain (no protocol):', def, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const domain = (resp.getResponseText() || def).trim().replace(/^https?:\/\//,'').replace(/\/+$/,'');
  PropertiesService.getUserProperties().setProperty('CANVAS_DOMAIN', domain);
  ui.alert('Saved: ' + domain);
}
function uiSetCanvasToken() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Paste your Canvas API Access Token (Account → Settings → New Access Token):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const token = (resp.getResponseText() || '').trim();
  if (!token) { ui.alert('No token entered.'); return; }
  PropertiesService.getUserProperties().setProperty('CANVAS_TOKEN', token);
  ui.alert('Token saved.');
}
function getCanvasBase_() {
  const domain = PropertiesService.getUserProperties().getProperty('CANVAS_DOMAIN') || 'canvas.gmu.edu';
  return 'https://' + domain + '/api/v1';
}
function getCanvasToken_() {
  const token = PropertiesService.getUserProperties().getProperty('CANVAS_TOKEN');
  if (!token) throw new Error('Canvas token not set. Use "AI + Canvas → Set Canvas API Token".');
  return token;
}
function parseNextLink_(linkHeader) {
  if (!linkHeader) return null;
  for (const part of linkHeader.split(',')) {
    const m = part.match(/<([^>]+)>;\s*rel="next"/i);
    if (m) return m[1];
  }
  return null;
}
function fetchCanvasJsonPaginated_(endpoint, params) {
  const base  = getCanvasBase_();
  const token = getCanvasToken_();
  const query = Object.keys(params || {}).map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`).join('&');
  let url = `${base}${endpoint}${query ? ('?' + query) : ''}`;
  let all = [];
  while (url) {
    const res = UrlFetchApp.fetch(url, { method:'get', headers:{ Authorization: 'Bearer ' + token }, muteHttpExceptions:true });
    if (res.getResponseCode() >= 400) throw new Error('Canvas error '+res.getResponseCode()+': '+res.getContentText());
    const page = JSON.parse(res.getContentText());
    if (Array.isArray(page)) all = all.concat(page); else all.push(page);
    url = parseNextLink_(res.getAllHeaders()['Link'] || res.getAllHeaders()['link']);
    Utilities.sleep(100);
  }
  return all;
}
function uiImportFromCanvasApi() {
  const ui = SpreadsheetApp.getUi();
  try { getCanvasToken_(); } catch (e) { ui.alert(e.message); return; }

  const rawDataSheet = getRawDataSheet();
  const dashboardSheet = getDashboardSheet();
  const courses = fetchCanvasJsonPaginated_('/courses', {
    enrollment_state: 'active',
    per_page: 100
  }).filter(c => !c.access_restricted_by_date);

  if (!courses.length) { ui.alert('No active courses found.'); return; }

  let rows = [];
  let totalAssignments = 0;
  let gradesFound = 0;

  for (const c of courses) {
    const courseLabel = (c.course_code || c.name || '').toString().trim();
    const courseId = c.id;
    
    // Get assignment groups with weights
    const assignmentGroups = fetchCanvasJsonPaginated_(`/courses/${courseId}/assignment_groups`, { per_page: 100 });
    const groupWeights = {};
    const groupNames = {};
    
    for (const group of assignmentGroups) {
      groupWeights[group.id] = group.group_weight || 0;
      groupNames[group.id] = group.name || '';
    }
    
    // Get assignments
    const assignments = fetchCanvasJsonPaginated_(`/courses/${courseId}/assignments`, { per_page: 100 });
    
    // Get current user ID for submissions
    let userId = null;
    try {
      const userInfo = fetchCanvasJsonPaginated_('/users/self', {});
      userId = userInfo[0]?.id;
    } catch (e) {
      console.log('Could not get user ID for submissions');
    }
    
    for (const a of assignments) {
      if (!a) continue;
      totalAssignments++;
      
      // Get assignment group info
      const groupId = a.assignment_group_id;
      const groupWeight = groupWeights[groupId] || '';
      const groupName = groupNames[groupId] || '';
      
      // Get submission info if user ID available
      let submission = null;
      if (userId && a.id) {
        try {
          const submissions = fetchCanvasJsonPaginated_(`/courses/${courseId}/assignments/${a.id}/submissions/${userId}`, {});
          submission = submissions[0];
          if (submission && submission.score != null) {
            gradesFound++;
          }
        } catch (e) {
          // Submission fetch failed, continue without submission data
        }
      }
      
      // Calculate days until due and overdue status
      let daysUntilDue = '';
      let isOverdue = false;
      if (a.due_at) {
        const dueDate = new Date(a.due_at);
        const today = new Date();
        const diffTime = dueDate - today;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        daysUntilDue = diffDays;
        isOverdue = diffDays < 0;
      }
      
      // Build comprehensive row with ALL Canvas assignment fields
      rows.push([
        // Core Assignment Fields
        a.id || '',
        a.name || '',
        stripHtml_(a.description || ''),
        a.course_id || '',
        a.assignment_group_id || '',
        a.position || '',
        
        // Dates & Times (keep as ISO strings for better compatibility)
        a.due_at || '',
        a.unlock_at || '',
        a.lock_at || '',
        a.created_at || '',
        a.updated_at || '',
        
        // Points & Grading
        a.points_possible != null ? a.points_possible : '',
        a.grading_type || '',
        a.grading_standard_id || '',
        a.omit_from_final_grade || false,
        
        // Submission Settings
        Array.isArray(a.submission_types) ? a.submission_types.join(', ') : (a.submission_types || ''),
        Array.isArray(a.allowed_extensions) ? a.allowed_extensions.join(', ') : (a.allowed_extensions || ''),
        a.allowed_attempts || '',
        a.annotatable_attachment_id || '',
        
        // Status & Visibility
        a.published || false,
        a.muted || false,
        a.anonymous_submissions || false,
        a.anonymous_grading || false,
        a.anonymous_instructor_annotations || false,
        a.hide_in_gradebook || false,
        a.post_to_sis || false,
        a.integration_id || '',
        a.integration_data || '',
        
        // Peer Reviews & Moderation
        a.peer_reviews || false,
        a.automatic_peer_reviews || false,
        a.peer_review_count || '',
        a.peer_reviews_assign_at || '',
        a.intra_group_peer_reviews || false,
        a.moderated_grading || false,
        a.grader_count || '',
        a.final_grader_id || '',
        a.grader_comments_visible_to_graders || false,
        a.graders_anonymous_to_graders || false,
        a.grader_names_visible_to_final_grader || false,
        
        // External Tools & Plagiarism
        a.turnitin_enabled || false,
        a.vericite_enabled || false,
        a.turnitin_settings ? JSON.stringify(a.turnitin_settings) : '',
        a.external_tool_tag_attributes ? JSON.stringify(a.external_tool_tag_attributes) : '',
        
        // Rubrics
        a.use_rubric_for_grading || false,
        a.rubric_settings ? JSON.stringify(a.rubric_settings) : '',
        a.rubric ? JSON.stringify(a.rubric) : '',
        a.free_form_criterion_comments || false,
        
        // URLs & Links
        a.html_url || '',
        a.submissions_download_url || '',
        a.quiz_id || '',
        a.discussion_topic ? JSON.stringify(a.discussion_topic) : '',
        
        // Assignment Group Info
        groupName,
        groupWeight,
        
        // Submission Status
        submission ? (submission.workflow_state || '') : '',
        submission ? (submission.score != null ? submission.score : '') : '',
        submission ? (submission.grade || '') : '',
        submission ? (submission.submitted_at || '') : '',
        submission ? (submission.graded_at || '') : '',
        submission ? (submission.late || false) : false,
        submission ? (submission.missing || false) : false,
        submission ? (submission.excused || false) : false,
        
        // Calculated Fields
        daysUntilDue,
        isOverdue
      ]);
    }
  }
  
  if (!rows.length) { ui.alert('No assignments found via API.'); return; }
  
  writeRowsInBatches_(rawDataSheet, rawDataSheet.getLastRow() + 1, rows, 300);
  
  // Update dashboard with formulas
  updateDashboardFormulas_(dashboardSheet);
  
  ui.alert(`Imported ${rows.length} assignments from Canvas API.\n` +
           `Found grades for ${gradesFound} assignments.\n` +
           `Assignment weights automatically applied.`);
}
