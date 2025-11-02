const fetch = require('node-fetch');
const XLSX = require('xlsx');
const sgMail = require('@sendgrid/mail');

// Configuration from GitHub Secrets
const FIREBASE_URL = process.env.FIREBASE_URL;
const SENDGRID_API_KEY = process.env.SENDGRID_API_KEY;
const RECIPIENT_EMAIL = process.env.RECIPIENT_EMAIL;
const SENDER_EMAIL = process.env.SENDER_EMAIL;

// Initialize SendGrid
sgMail.setApiKey(SENDGRID_API_KEY);

/**
 * Generate Statistics Excel file with yearly summary
 */
function generateStatsExcel(items) {
  console.log('📊 Generating yearly statistics Excel...');
  
  const data = items.map(item => {
    const cleanings = item.activities ? item.activities.filter(a => a.cleaning).length : 0;
    const inspections = item.activities ? item.activities.filter(a => a.inspection).length : 0;
    
    // Calculate activities for current year
    const currentYear = new Date().getFullYear();
    const yearlyCleanings = item.activities ? 
      item.activities.filter(a => a.cleaning && new Date(a.date).getFullYear() === currentYear).length : 0;
    const yearlyInspections = item.activities ? 
      item.activities.filter(a => a.inspection && new Date(a.date).getFullYear() === currentYear).length : 0;
    
    return {
      'Firefighter Name': item.name,
      'Item Type': item.type,
      'Serial': item.serial,
      'Model': item.model || 'N/A',
      'Cleanings (This Year)': yearlyCleanings,
      'Inspections (This Year)': yearlyInspections,
      'Total Cleanings (All Time)': cleanings,
      'Total Inspections (All Time)': inspections,
      'Total Activities (This Year)': yearlyCleanings + yearlyInspections,
      'Total Activities (All Time)': cleanings + inspections
    };
  });
  
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Yearly Statistics');
  
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

/**
 * Generate Activity Logs Excel file for the current year
 */
function generateLogsExcel(items) {
  console.log('📋 Generating yearly activity logs Excel...');
  
  const currentYear = new Date().getFullYear();
  const data = [];
  
  items.forEach(item => {
    if (item.activities && item.activities.length > 0) {
      item.activities.forEach(activity => {
        const activityDate = new Date(activity.date);
        // Only include activities from current year
        if (activityDate.getFullYear() === currentYear) {
          data.push({
            'Firefighter Name': activity.snapshotName || item.name,
            'Item Type': activity.snapshotType || item.type,
            'Serial Number': activity.snapshotSerial || item.serial,
            'Model': activity.snapshotModel || item.model || 'N/A',
            'Date': activity.date,
            'Advanced Cleaning': activity.cleaning ? 'Yes' : 'No',
            'Advanced Inspection': activity.inspection ? 'Yes' : 'No'
          });
        }
      });
    }
  });
  
  // Sort by date (most recent first)
  data.sort((a, b) => new Date(b.Date) - new Date(a.Date));
  
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `${currentYear} Activity Logs`);
  
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

/**
 * Main function - Fetch data, generate reports, and send email
 */
async function generateAndSendReport() {
  try {
    console.log('🚀 Starting yearly report generation...');
    console.log(`📅 Date: ${new Date().toLocaleString()}`);
    
    // Validate environment variables
    if (!FIREBASE_URL || !SENDGRID_API_KEY || !RECIPIENT_EMAIL || !SENDER_EMAIL) {
      throw new Error('Missing required environment variables. Please check GitHub Secrets.');
    }
    
    // Fetch data from Firebase
    console.log('🔍 Fetching data from Firebase...');
    const response = await fetch(FIREBASE_URL);
    
    if (!response.ok) {
      throw new Error(`Firebase request failed: ${response.status} ${response.statusText}`);
    }
    
    const items = await response.json() || [];
    console.log(`✅ Found ${items.length} items in database`);
    
    if (items.length === 0) {
      console.log('⚠️  No items found - sending notification email');
      
      await sgMail.send({
        to: RECIPIENT_EMAIL,
        from: SENDER_EMAIL,
        subject: 'Annual PPE Report - No Data',
        text: 'The yearly report was triggered but no items were found in the database.',
        html: '<h2>Annual PPE Report</h2><p>No items found in the database.</p>'
      });
      
      console.log('✅ Notification sent');
      return;
    }
    
    // Generate Excel files
    const statsBuffer = generateStatsExcel(items);
    const logsBuffer = generateLogsExcel(items);
    
    // Calculate statistics for email
    const currentYear = new Date().getFullYear();
    let totalCleanings = 0;
    let totalInspections = 0;
    let yearlyCleanings = 0;
    let yearlyInspections = 0;
    
    items.forEach(item => {
      if (item.activities) {
        totalCleanings += item.activities.filter(a => a.cleaning).length;
        totalInspections += item.activities.filter(a => a.inspection).length;
        yearlyCleanings += item.activities.filter(a => 
          a.cleaning && new Date(a.date).getFullYear() === currentYear
        ).length;
        yearlyInspections += item.activities.filter(a => 
          a.inspection && new Date(a.date).getFullYear() === currentYear
        ).length;
      }
    });
    
    // Prepare email
    const today = new Date().toLocaleDateString();
    const msg = {
      to: RECIPIENT_EMAIL,
      from: SENDER_EMAIL,
      subject: `Annual PPE Report ${currentYear} - ${today}`,
      text: `Please find attached the annual PPE statistics and activity logs for ${currentYear}.`,
      html: `
        <!DOCTYPE html>
        <html>
        <head>
          <style>
            body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
            .container { max-width: 600px; margin: 0 auto; padding: 20px; }
            h2 { color: #2563eb; border-bottom: 3px solid #2563eb; padding-bottom: 10px; }
            .stats { background: #f3f4f6; padding: 15px; border-radius: 8px; margin: 20px 0; }
            .stat-item { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #d1d5db; }
            .stat-item:last-child { border-bottom: none; }
            .stat-label { font-weight: bold; }
            .stat-value { color: #2563eb; font-weight: bold; }
            .year-highlight { background: #dbeafe; padding: 10px; border-radius: 5px; margin: 15px 0; text-align: center; font-size: 18px; font-weight: bold; color: #1e40af; }
            ul { background: #fff; padding: 15px 15px 15px 35px; border-left: 4px solid #10b981; }
            .footer { margin-top: 30px; padding-top: 20px; border-top: 2px solid #e5e7eb; font-size: 12px; color: #6b7280; }
          </style>
        </head>
        <body>
          <div class="container">
            <h2>📊 Annual PPE Report</h2>
            
            <div class="year-highlight">Year ${currentYear} Summary</div>
            
            <p>This is your automated annual report for the PPE Tracking System.</p>
            
            <div class="stats">
              <div class="stat-item">
                <span class="stat-label">Report Date:</span>
                <span class="stat-value">${today}</span>
              </div>
              <div class="stat-item">
                <span class="stat-label">Total Items:</span>
                <span class="stat-value">${items.length}</span>
              </div>
              <div class="stat-item">
                <span class="stat-label">Cleanings (${currentYear}):</span>
                <span class="stat-value">${yearlyCleanings}</span>
              </div>
              <div class="stat-item">
                <span class="stat-label">Inspections (${currentYear}):</span>
                <span class="stat-value">${yearlyInspections}</span>
              </div>
              <div class="stat-item">
                <span class="stat-label">Total Cleanings (All Time):</span>
                <span class="stat-value">${totalCleanings}</span>
              </div>
              <div class="stat-item">
                <span class="stat-label">Total Inspections (All Time):</span>
                <span class="stat-value">${totalInspections}</span>
              </div>
            </div>
            
            <h3>📎 Attached Files</h3>
            <ul>
              <li><strong>PPE_Annual_Statistics_${currentYear}.xlsx</strong> - Summary statistics including current year and all-time totals</li>
              <li><strong>PPE_Activity_Logs_${currentYear}.xlsx</strong> - Detailed activity logs for all events in ${currentYear}</li>
            </ul>
            
            <div class="footer">
              <p><em>This is an automated annual message from the PPE Tracking System.</em></p>
              <p>Generated by GitHub Actions on ${new Date().toLocaleString()}</p>
            </div>
          </div>
        </body>
        </html>
      `,
      attachments: [
        {
          content: statsBuffer.toString('base64'),
          filename: `PPE_Annual_Statistics_${currentYear}.xlsx`,
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          disposition: 'attachment'
        },
        {
          content: logsBuffer.toString('base64'),
          filename: `PPE_Activity_Logs_${currentYear}.xlsx`,
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          disposition: 'attachment'
        }
      ]
    };
    
    // Send email
    console.log('📧 Sending email to:', RECIPIENT_EMAIL);
    await sgMail.send(msg);
    console.log('✅ Annual report sent successfully!');
    console.log(`📬 Email sent with 2 Excel attachments for year ${currentYear}`);
    
  } catch (error) {
    console.error('❌ Error generating or sending report:', error);
    
    // Try to send error notification
    try {
      await sgMail.send({
        to: RECIPIENT_EMAIL,
        from: SENDER_EMAIL,
        subject: 'ERROR: Annual PPE Report Failed',
        text: `The annual report failed to generate.\n\nError: ${error.message}\n\nPlease check the GitHub Actions logs for more details.`,
        html: `
          <h2 style="color: #dc2626;">⚠️ Annual Report Generation Failed</h2>
          <p><strong>Error:</strong> ${error.message}</p>
          <p>Please check the <a href="https://github.com/${process.env.GITHUB_REPOSITORY}/actions">GitHub Actions logs</a> for more details.</p>
          <p><em>Timestamp: ${new Date().toLocaleString()}</em></p>
        `
      });
      console.log('📧 Error notification sent');
    } catch (emailError) {
      console.error('❌ Failed to send error notification:', emailError);
    }
    
    // Exit with error code so GitHub Actions marks it as failed
    process.exit(1);
  }
}

// Run the main function
generateAndSendReport();
