const cron = require('node-cron');
const invoiceReportMTD_CH = require('../services/DatabaseService');

function startScheduler() {
    
    cron.schedule('* * * * *', async () => {
        console.log('Running a task every minute');
        await invoiceReportMTD_CH(); 
    });

   
}

module.exports = startScheduler;
