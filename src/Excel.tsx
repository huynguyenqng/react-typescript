import React from "react";
import axios from "axios";
const ExcelJS = require("exceljs");

interface IProps {}
//
class Excel extends React.Component<IProps> {
    constructor(props: any) {
        super(props);
        this.state = { data: "" };
    }
    componentDidMount() { }
    async clickHandler(e: any) {
        console.log("------clickHandler------");
        e.preventDefault();

        var wb = new ExcelJS.Workbook();
        wb.xlsx.readFile('./test.xlsx');
        console.log("-----------", wb)


        const res = await axios.get("./test.xlsx", { responseType: "arraybuffer" });
        const data = new Uint8Array(res.data);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(data);
        //data
        const items = [
            { id: 1001, name: "みかん", price: 180 },
            { id: 1002, name: "りんご", price: 210 },
            { id: 1003, name: "バナナ", price: 170 },
        ];
        console.log(items);
        const worksheet = workbook.getWorksheet("sheet1");
        worksheet.pageSetup = { orientation: "portrait" };
        const startRow = 4;
        let iCount = 0;
        let row = worksheet.getRow(1);
        for (const item of items) {
            let pos = startRow + iCount;
            row = worksheet.getRow(pos);
            console.log(item);
            row.getCell(1).value = item.id;
            row.getCell(2).value = item.name;
            row.getCell(3).value = item.price;
            iCount += 1;
        }
        //save
        const uint8Array = await workbook.xlsx.writeBuffer();
        //console.log(uint8Array);
        const blob = new Blob([uint8Array], { type: "application/octet-binary" });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `out.xlsx`;
        a.click();
        a.remove();
    }
    render() {
        return (
            <div>
                <h1>xls6: read templete</h1>
                <hr />
                <button onClick={(e) => this.clickHandler(e)}>Read</button>
            </div>
        );
    }
}

export default Excel;
