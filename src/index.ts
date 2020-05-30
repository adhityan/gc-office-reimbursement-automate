import { Command, flags } from '@oclif/command';
import * as stringSimilarity from 'string-similarity';
import * as moment from 'moment-timezone';
import * as prompts from 'prompts';
import * as chalk from 'chalk';
import * as xlsx from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';

import { sourceFormat } from './types';

class Reimbursements extends Command {
    static description = 'describe the command here';

    static flags = {
        version: flags.version({ char: 'v' }),
        help: flags.help({ char: 'h' })
    };

    static args = [
        {
            name: 'input',
            required: true,
            description: 'The citibank file to read'
        },
        {
            name: 'output',
            required: true,
            description: 'The current work in progress reimbursements file'
        }
    ];

    originalLength: number = 0;
    outputFileName: string | undefined;
    primaryOutputWorkbook: xlsx.WorkBook | undefined;

    async writeFile(workbook: xlsx.WorkBook | null = null, filename: string | null = null) {
        if (!workbook) workbook = <xlsx.WorkBook>this.primaryOutputWorkbook;
        if (!filename) filename = <string>this.outputFileName;

        await xlsx.writeFile(workbook, filename);
    }

    async initPrimaryOutput() {
        this.primaryOutputWorkbook = xlsx.utils.book_new();

        xlsx.utils.book_append_sheet(this.primaryOutputWorkbook, xlsx.utils.aoa_to_sheet([]), 'Candidates');
        xlsx.utils.book_append_sheet(this.primaryOutputWorkbook, xlsx.utils.aoa_to_sheet([]), 'Finalized');
        xlsx.utils.book_append_sheet(this.primaryOutputWorkbook, xlsx.utils.aoa_to_sheet([]), 'Discarded');

        await this.writeFile();
    }

    async initCandidatesSheet(primaryInputSheet: xlsx.WorkSheet): Promise<sourceFormat[]> {
        moment.tz.setDefault('Asia/Dubai');

        const rawData: (string | number)[][] = xlsx.utils.sheet_to_json(primaryInputSheet, {
            header: 1,
            raw: false
        });

        const data = rawData
            .filter((row) => row.length === 6)
            .map((row) => [row[0], row[1], row[2], row[3], '', row[5]])
            .reverse();

        const workSheet = <xlsx.WorkSheet>this.primaryOutputWorkbook?.Sheets[this.primaryOutputWorkbook.SheetNames[0]];

        xlsx.utils.sheet_add_aoa(workSheet, data, {
            origin: 0
        });
        await this.writeFile();

        this.originalLength = data.length;
        return data.map(
            (row) =>
                <sourceFormat>{
                    date: <string>row[0],
                    //date: moment(<string>row[0], 'DD/MM/YYYY').toDate(),
                    name: row[1],
                    cost: -1 * <number>row[2],
                    currency: row[3],
                    id: (<string>row[5]).replace(/'/g, '')
                }
        );
    }

    async recreateCandidatesSheet(inputData: sourceFormat[]) {
        const data = inputData.map((row) => [row.date, row.name, -1 * row.cost, row.currency, '', row.id]);

        const workSheet = <xlsx.WorkSheet>this.primaryOutputWorkbook?.Sheets[this.primaryOutputWorkbook.SheetNames[0]];
        const rows = <number>workSheet['!rows']?.length;
        const diff = rows - data.length;

        for (let i = 0; i <= diff; i++) {
            data.push(['', '', '', '', '', '']);
        }

        xlsx.utils.sheet_add_aoa(workSheet, data, {
            origin: 0
        });
        await this.writeFile();
    }

    async reverseCandidatesSheet(inputData: sourceFormat[]) {
        const data = inputData.map((row) => [row.date, row.name, -1 * row.cost, row.currency, '', row.id]).reverse();

        const workSheet = <xlsx.WorkSheet>this.primaryOutputWorkbook?.Sheets[this.primaryOutputWorkbook.SheetNames[0]];

        data.length = this.originalLength;
        xlsx.utils.sheet_add_aoa(workSheet, data, {
            origin: 0
        });
        await this.writeFile();
    }

    async addToDiscardedPile(row: sourceFormat) {
        const data = [row.date, row.name, -1 * row.cost, row.currency, '', row.id];

        const workSheet = <xlsx.WorkSheet>this.primaryOutputWorkbook?.Sheets[this.primaryOutputWorkbook.SheetNames[2]];

        xlsx.utils.sheet_add_aoa(workSheet, [data], {
            origin: -1
        });
        await this.writeFile();
    }

    async addEntryToOutputSheet(entry: sourceFormat, title: string, type: string) {
        const record = [entry.date, title, type, entry.cost];
        const workSheet = <xlsx.WorkSheet>this.primaryOutputWorkbook?.Sheets[this.primaryOutputWorkbook.SheetNames[1]];

        xlsx.utils.sheet_add_aoa(workSheet, [record], {
            origin: -1
        });

        await this.writeFile();
    }

    async sortOutputSheet() {
        const workSheet = <xlsx.WorkSheet>this.primaryOutputWorkbook?.Sheets[this.primaryOutputWorkbook.SheetNames[1]];

        const rawData: (string | number)[][] = xlsx.utils.sheet_to_json(workSheet, {
            header: 1,
            raw: false
        });

        let data = rawData
            .filter((row) => row.length === 4)
            .map((row) => ({
                rawDate: <string>row[0],
                date: moment(<string>row[0], 'DD/MM/YYYY').toDate(),
                title: row[1],
                type: row[2],
                cost: <number>row[3]
            }));

        const result = data
            .sort((a, b) => +b.date - +a.date)
            .map((row) => [`x${row.rawDate}`, row.title, row.type, row.cost]);

        xlsx.utils.sheet_add_aoa(workSheet, result.reverse(), {
            origin: 0
        });

        await this.writeFile();
    }

    async run() {
        const { args } = this.parse(Reimbursements);

        const inputFile = path.resolve(args.input);
        if (!fs.existsSync(inputFile)) {
            this.error(`Input file '${inputFile}' not found`);
        }

        this.outputFileName = path.resolve(args.output);
        if (fs.existsSync(this.outputFileName)) {
            this.error(`Output file '${this.outputFileName}' already exists`);
        }
        this.log(`Creating output file '${this.outputFileName}'`);
        await this.initPrimaryOutput();

        this.log(`Reading input file '${inputFile}'`);
        const inputBook = xlsx.readFile(inputFile);
        const primaryInputSheet = inputBook.Sheets[inputBook.SheetNames[0]];
        const data = await this.initCandidatesSheet(primaryInputSheet);

        while (data.length > 0) {
            this.log(chalk.gray(`${data.length} candidates remaining`));

            const sourceEntry = data[0];
            data.splice(0, 1);

            const firstWord = sourceEntry.name.split(' ')[0];
            const similarEntries = data.filter((entry) => {
                const similarityScore = stringSimilarity.compareTwoStrings(entry.name, sourceEntry.name);
                if (entry.name.includes(firstWord) && similarityScore > 0.8) return true;
                else return false;
            });

            this.log(chalk.cyan('similarEntries'), similarEntries);
            this.log(chalk.redBright('thisEntry'), sourceEntry);

            const { doAdd } = <{ doAdd: boolean }>await prompts({
                type: 'confirm',
                name: 'doAdd',
                message: `Do you want to add this entry?`,
                initial: false
            });

            if (doAdd) {
                const { title } = <{ title: string }>await prompts({
                    type: 'text',
                    name: 'title',
                    message: `What is the name of this entry?`
                });

                const { type } = <{ type: string }>await prompts({
                    type: 'select',
                    name: 'type',
                    message: `What is the type of this record?`,
                    choices: [
                        { title: 'Team Meals', value: 'Team Meals' },
                        { title: 'Training Fees', value: 'Training Fees' },
                        { title: 'Office Supplies', value: 'Office Supplies' },
                        { title: 'Software', value: 'Software' },
                        { title: 'Telephones', value: 'Telephones' },
                        { title: 'Travel', value: 'Travel' },
                        { title: 'Other', value: 'Other' },
                        { title: 'Postage', value: 'Postage' },
                        { title: 'Stationery', value: 'Stationery' },
                        { title: 'Subscriptions', value: 'Subscriptions' }
                    ],
                    initial: 3
                });

                await this.addEntryToOutputSheet(sourceEntry, title, type);

                if (similarEntries.length > 0) {
                    const { addSimilar } = <{ addSimilar: boolean }>await prompts({
                        type: 'confirm',
                        name: 'addSimilar',
                        message: `Do you want to add the ${similarEntries.length} similar entries found?`,
                        initial: false
                    });

                    if (addSimilar) {
                        for (let i = 0; i < similarEntries.length; i++) {
                            const element = similarEntries[i];
                            await this.addEntryToOutputSheet(element, title, type);

                            const index = data.findIndex((entry) => {
                                if (entry.name === element.name) return true;
                                else return false;
                            });
                            data.splice(index, 1);
                        }
                    }
                }
            } else {
                await this.addToDiscardedPile(sourceEntry);

                if (similarEntries.length > 0) {
                    const { removeSimilar } = <{ removeSimilar: boolean }>await prompts({
                        type: 'confirm',
                        name: 'removeSimilar',
                        message: `Do you want to remove the ${similarEntries.length} similar entries found?`,
                        initial: false
                    });

                    if (removeSimilar) {
                        for (let i = 0; i < similarEntries.length; i++) {
                            const element = similarEntries[i];
                            await this.addToDiscardedPile(element);

                            const index = data.findIndex((entry) => {
                                if (entry.name === element.name) return true;
                                else return false;
                            });
                            data.splice(index, 1);
                        }
                    }
                }
            }

            await this.recreateCandidatesSheet(data);
            this.log('\n\n\n\n\n');
            this.log(
                chalk.yellowBright(
                    '------------------------------------------------------------------------------------------------------------------------------------------------------------'
                )
            );
            this.log('\n\n\n\n\n');

            const { doTerminate } = <{ doTerminate: boolean }>await prompts({
                type: 'confirm',
                name: 'doTerminate',
                message: `Do you want to terminate?`,
                initial: false
            });

            if (doTerminate) break;
        }

        this.log('\n');
        await this.reverseCandidatesSheet(data);
        await this.sortOutputSheet();
        this.log(chalk.greenBright('Done!'));
    }
}

export = Reimbursements;
