const readXlsxFile = require('read-excel-file/node')
const ObjectsToCsv = require('objects-to-csv');
const path = require('path');
const fs = require('fs');

const FILE = '/Users/rasmuslind/Desktop/report.xlsx'

let formatTicket = (row, heading, linkToIssue = true) => {
    let ticketData = {}

    ticketData.priority = row[1] // Use Serverity column for priority in Jira ticket
    ticketData.labels = row[5] // Use Page column for label in Jira ticket
    ticketData.issueType = "bug"
    ticketData.summary = row[6] // Use Summary column for summary in Jira ticket

    // Very specific to BoostEvents as we wanted to link to these tickets to other issues. 
    // To link to any other issue create another issue in jira and set the jira name in ticket.blocks to link the two
    // To disable, just change function parameter linkToIssue to false and then the ticket will not contain a link to other tickets
    if(linkToIssue) {
        if(row[7].includes("portal.voiceboxer")) {
            ticketData.blocks = "VPAT-824"
        }
        if(row[7].includes("event.voiceboxer")) {
            ticketData.blocks = "VPAT-825"
        }
    }

    // Format a string for description in Jira from the other rows in the excel report
    let description = `h3. ${heading[2]}\n${row[2]}\n\nh3. ${heading[3]}\n${row[3]}\n\nh3. ${heading[4]}\n${row[4]}\n\nh3. ${heading[5]}\n${row[5]}\n\nh3. ${heading[7]}\n${row[7]}\n\nh3. ${heading[8]}\n${row[8]}\n\nh3. ${heading[9]}\n${row[9]}\n\nh3. ${heading[10]}\n${row[10]}\n\nh3. ${heading[11]}\n${row[11]}\n\nh3. ${heading[12]}\n${row[12]}\n\nh3. ${heading[13]}\n${row[13]}`
    
    ticketData.description = description

    return ticketData
}

const convertToCSV = async (ticketsArray, sheet) => {
    const csv = new ObjectsToCsv(ticketsArray);

    const csvOutputDir = './csv-output'

    if(!fs.existsSync(csvOutputDir)) {
        fs.mkdir(path.join(__dirname, 'csv-output'), (err) => {
            if (err) {
                return console.error(err);
            }
            console.log('Directory created successfully!');
        });
    } 
    
    await csv.toDisk(`./csv-output/${sheet}.csv`);
}

const convertToTicketFromFile = (file, sheet = 1) => {

    let tickets = []

    return readXlsxFile(file, { sheet }).then((rows) => {

        rows.forEach((row) => {
            if(row[0] && row[0].includes("VOI")) {

                let ticketData = formatTicket(row, rows[2])
                
                tickets.push(ticketData)
            }
        })

        return convertToCSV(tickets, sheet) 

    })
}

convertToTicketFromFile(FILE, "Axe")
