import pkg from 'xlsx';
const { utils, writeFile } = pkg;
import XlsxStreamReader from 'xlsx-stream-reader';
import { createReadStream } from 'fs';
import { promises as fs } from 'fs';
import path from 'path';
import yargs from 'yargs';
import { hideBin } from 'yargs/helpers';

const argv = yargs(hideBin(process.argv))
    .option('input', {
        alias: 'i',
        description: 'Input XLSX file path',
        type: 'string',
        demandOption: true
    })
    .option('rows', {
        alias: 'r',
        description: 'Number of rows per split file',
        type: 'number',
        default: 1000
    })
    .help()
    .argv;

const writeChunkToFile = async (chunk, sheetName, index, baseFileName, outputDir) => {
    const newWorkbook = utils.book_new();
    const newWorksheet = utils.aoa_to_sheet(chunk, {
        raw: true,
        cellDates: true
    });
    
    utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
    
    const outputFile = path.join(outputDir, `${baseFileName}_part${index + 1}.xlsx`);
    writeFile(newWorkbook, outputFile);
    console.log(`Created: ${outputFile} (${chunk.length} rows)`);
};

const splitXlsx = async (inputFile, rowsPerFile = 1000) => {
    try {
        console.log('Reading file:', inputFile);
        
        // Create output directory
        const outputDir = path.join(process.cwd(), 'output');
        await fs.mkdir(outputDir, { recursive: true });
        
        const inputFileName = path.basename(inputFile, path.extname(inputFile));
        
        return new Promise((resolve, reject) => {
            const workBookReader = new XlsxStreamReader();
            let currentChunk = [];
            let fileIndex = 0;
            let totalRows = 0;
            let headers = null;
            
            workBookReader.on('error', (error) => {
                console.error('Error reading spreadsheet:', error);
                reject(error);
            });

            workBookReader.on('worksheet', (workSheetReader) => {
                console.log(`Processing worksheet: ${workSheetReader.name}`);
                
                workSheetReader.on('row', async (row) => {
                    if (!headers) {
                        headers = row.values.slice(1); // Skip empty first element
                        console.log('Headers:', headers);
                        return;
                    }
                    
                    // Add row to current chunk
                    const rowValues = row.values.slice(1); // Skip empty first element
                    if (rowValues.some(cell => cell !== null && cell !== undefined)) {
                        currentChunk.push(rowValues);
                        totalRows++;
                        
                        // If chunk is complete, write it to file
                        if (currentChunk.length >= rowsPerFile) {
                            try {
                                await writeChunkToFile(
                                    [headers, ...currentChunk],
                                    workSheetReader.name,
                                    fileIndex,
                                    inputFileName,
                                    outputDir
                                );
                                fileIndex++;
                                currentChunk = [];
                            } catch (error) {
                                console.error('Error writing chunk:', error);
                                workSheetReader.abort();
                                reject(error);
                            }
                        }
                    }
                });

                workSheetReader.on('end', async () => {
                    // Write remaining rows
                    if (currentChunk.length > 0) {
                        try {
                            await writeChunkToFile(
                                [headers, ...currentChunk],
                                workSheetReader.name,
                                fileIndex,
                                inputFileName,
                                outputDir
                            );
                        } catch (error) {
                            console.error('Error writing final chunk:', error);
                            reject(error);
                            return;
                        }
                    }
                    console.log(`Processed ${totalRows} total rows`);
                });

                // Process the worksheet
                workSheetReader.process();
            });

            workBookReader.on('end', () => {
                console.log('Finished processing workbook');
                resolve();
            });

            // Create read stream and pipe it to the reader
            const fileStream = createReadStream(inputFile);
            fileStream.pipe(workBookReader);
        });
    } catch (error) {
        console.error('Error splitting the XLSX file:', error);
        console.error('Stack trace:', error.stack);
        process.exit(1);
    }
};

// Execute the split operation
splitXlsx(argv.input, argv.rows);
