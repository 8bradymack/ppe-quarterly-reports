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
 * Generate Statistics Excel file
 */
function generateStatsExcel(items) {
  console.log('📊 Generating statistics Excel...');
  
  const data = items.map(item => {
    const cleanings = item.activities ? item.activities.filter(a => a.cleaning).length : 0;
    const inspections = item.activities ? item.activities.filter(a => a.inspection).length : 0;
    
    return {
      'Firefighter Name': item.name,
      'Item Type': item.type,
      'Serial': item.serial,
      'Model': item.model || 'N/A',
      'Total Cleanings': cleanings,
      'Total Inspections': inspections,
      'Total Activities': cleanings + inspections
    };
  });
  
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Statistics');
  
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

/**
 * Generate Activity Logs Excel file
 */
function generateLogsExcel(items) {
  console.log('📋 Generating activity logs Excel...');
  
  const data = [];
  
  items.forEach(item => {
    if (item.activities && item.activities.length > 0) {
      item.activities.forEach(activity => {
        data.push({
          'Firefighter Name': activity.snapshotName || item.name,
          'Item Type': activity.snapshotType || item.type,
          'Serial Number': activity.snapshotSerial || item.serial,
          'Model': activity.snapshotModel || item.model || 'N/A',
          'Date': activity.date,
          'Advanced Cleaning': activity.cleaning ? 'Yes' : 'No',
          'Advanced Inspection': activity.inspection ? 'Yes' : 'No'
        });
      });
    } else {
      // Include items with no activities
      data.push({
        'Firefighter Name': item.name,
        'Item Type': item.type,
        'Serial Number': item.serial,
        'Model': item.model || 'N/A',
        'Date': 'No activities',
        'Advanced Cleaning': 'No',
        'Advanced Inspection': 'No'
      });
    }
  });
  
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Activity Logs');
  
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

/**
 * Main function - Fetch data, generate reports, and send email
 */
async function generateAndSendReport() {
  try {
    console.log('🚀 Starting quarterly report generation...');
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
        subject: 'Quarterly PPE Report - No Data',
        text: 'The quarterly report was triggered but no items were found in the database.',
        html: '<h2>Quarterly PPE Report</h2><p>No items found in the database.</p>'
      });
      
      console.log('✅ Notification sent');
      return;
    }
    
    // Generate Excel files
    const statsBuffer = generateStatsExcel(items);
    const logsBuffer = generateLogsExcel(items);
    
    // Calculate statistics for email
    let totalCleanings = 0;
    let totalInspections = 0;
    items.forEach(item => {
      if (item.activities) {
        totalCleanings += item.activities.filter(a => a.cleaning).length;
        totalInspections += item.activities.filter(a => a.inspection).length;
      }
    });
    
    // Prepare email
    const today = new Date().toLocaleDateString();
    const msg = {
      to: RECIPIENT_EMAIL,
      from: SENDER_EMAIL,
      subject: `Quarterly PPE Report - ${today}`,
      text: `Please find attached the quarterly PPE statistics and activity logs for ${today}.`,
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
            ul { background: #fff; padding: 15px 15px 15px 35px; border-left: 4px solid #10b981; }
            .footer { margin-top: 30px; padding-top: 20px; border-top: 2px solid #e5e7eb; font-size: 12px; color: #6b7280; }
          </style>
        </head>
        <body>
          <div class="container">
            <h2>📊 Quarterly PPE Report</h2>
            
            <p>This is your automated quarterly report for the PPE Tracking System.</p>
            
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
                <span class="stat-label">Total Cleanings:</span>
                <span class="stat-value">${totalCleanings}</span>
              </div>
              <div class="stat-item">
                <span class="stat-label">Total Inspections:</span>
                <span class="stat-value">${totalInspections}</span>
              </div>
            </div>
            
            <h3>📎 Attached Files</h3>
            <ul>
              <li><strong>PPE_Statistics.xlsx</strong> - Summary statistics for all items</li>
              <li><strong>PPE_Activity_Logs.xlsx</strong> - Detailed activity logs for all recorded events</li>
            </ul>
            
            <div class="footer">
              <p><em>This is an automated message from the PPE Tracking System.</em></p>
              <p>Generated by GitHub Actions on ${new Date().toLocaleString()}</p>
            </div>
          </div>
        </body>
        </html>
      `,
      attachments: [
        {
          content: statsBuffer.toString('base64'),
          filename: `PPE_Statistics_${today.replace(/\//g, '-')}.xlsx`,
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          disposition: 'attachment'
        },
        {
          content: logsBuffer.toString('base64'),
          filename: `PPE_Activity_Logs_${today.replace(/\//g, '-')}.xlsx`,
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          disposition: 'attachment'
        }
      ]
    };
    
    // Send email
    console.log('📧 Sending email to:', RECIPIENT_EMAIL);
    await sgMail.send(msg);
    console.log('✅ Quarterly report sent successfully!');
    console.log('📬 Email sent with 2 Excel attachments');
    
  } catch (error) {
    console.error('❌ Error generating or sending report:', error);
    
    // Try to send error notification
    try {
      await sgMail.send({
        to: RECIPIENT_EMAIL,
        from: SENDER_EMAIL,
        subject: 'ERROR: Quarterly PPE Report Failed',
        text: `The quarterly report failed to generate.\n\nError: ${error.message}\n\nPlease check the GitHub Actions logs for more details.`,
        html: `
          <h2 style="color: #dc2626;">⚠️ Quarterly Report Generation Failed</h2>
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
