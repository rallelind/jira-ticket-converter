const readXlsxFile = require('read-excel-file/node')
const ObjectsToCsv = require('objects-to-csv');
const path = require('path');
const fs = require('fs');

const FILE = '/example/example.xlsx'

let formatTicket = (row, heading) => {
    let ticketData = {}

    let description = `${heading[2]}\n${row[2]}\n${heading[3]}\n${row[3]}\n${heading[4]}\n${row[4]}\n${heading[5]}\n${row[5]}\n${heading[7]}\n${row[7]}\n${heading[8]}\n${row[8]}\n${heading[9]}\n${row[9]}\n${heading[10]}\n${row[10]}\n${heading[11]}\n${row[11]}\n${heading[12]}\n${row[12]}\n${heading[13]}\n${row[13]}`

    ticketData.priority = row[1]
    ticketData.issueType = "bug"
    ticketData.summary = row[6]
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

convertToTicketFromFile(FILE, "Color Contrast")
