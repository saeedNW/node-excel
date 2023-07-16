/** import express */
const express = require("express");
/** create app */
const app = express();

app.get('/', (req, res) => {
    res.json({
        message: "go to one of the following routes",
        routes: {
            "read XLS file": "http://localhost:3000/read-xls",
            "write XLS file": "http://localhost:3000/write-xls",
            "read XLSX file": "http://localhost:3000/read-xlsx",
            "write XLSX file": "http://localhost:3000/write-xlsx",
        }
    });
});

app.get('/read-xls', (req, res) => {
    /** import xlsx module */
    const xlsx = require('xlsx');

    /** read example file */
    const file = xlsx.readFile('./example.xls');

    let data = [];

    /** read file's sheet names */
    const sheets = file.SheetNames;

    for (const sheet of sheets) {
        /** convert sheets' data to json */
        const temp = xlsx.utils.sheet_to_json(file.Sheets[sheets]);

        temp.forEach((res) => {
            data.push(res)
        })
    }

    /** return data */
    return res.json({data});
});

app.get("/write-xls", (req, res) => {
    /** import xlsx module */
    const xlsx = require('xlsx');
    /** import mkdir module */
    const mkdir = require("mkdir");

    /** create an empty workbook */
    const workbook = xlsx.utils.book_new();

    /** create worksheet data */
    const worksheet = xlsx.utils.json_to_sheet([
        {name: "saeed", age: "25"},
        {name: "saber", age: "24"},
    ]);

    /** add worksheet data to workbook */
    xlsx.utils.book_append_sheet(workbook, worksheet, 'users');

    /** define file name */
    const file = `./new_files/users_${new Date().getTime()}.xls`;

    mkdir.mkdirsSync("./new_files");

    /** create a xls file from data */
    xlsx.writeFile(workbook, file);

    /** download the file */
    res.download(file);
})

app.get('/read-xlsx', (req, res) => {
    /** import xlsx module */
    const xlsx = require('xlsx');

    /** read example file */
    const file = xlsx.readFile('./example.xlsx');

    let data = [];

    const sheets = file.SheetNames;

    for (const sheet of sheets) {
        const temp = xlsx.utils.sheet_to_json(file.Sheets[sheets]);

        temp.forEach((res) => {
            data.push(res)
        })
    }

    /** return data */
    return res.json({data});
});

app.get("/write-xlsx", (req, res) => {
    /** import xlsx module */
    const xlsx = require('xlsx');
    /** import mkdir module */
    const mkdir = require("mkdir");

    /** create an empty workbook */
    const workbook = xlsx.utils.book_new();

    /** create worksheet data */
    const worksheet = xlsx.utils.json_to_sheet([
        {name: "saeed", age: "25"},
        {name: "saber", age: "24"},
    ]);

    /** add worksheet data to workbook */
    xlsx.utils.book_append_sheet(workbook, worksheet, 'users');

    /** define file name */
    const file = `./new_files/users_${new Date().getTime()}.xlsx`;

    mkdir.mkdirsSync("./new_files");

    /** create a xls file from data */
    xlsx.writeFile(workbook, file);

    /** download the file */
    res.download(file);
})

/** starting server */
app.listen(3000, () => {
    console.log("Running");
})