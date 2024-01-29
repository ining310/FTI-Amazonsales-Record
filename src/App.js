import { useState } from "react";
import * as XLSX from "xlsx-js-style";
import 'mdb-react-ui-kit/dist/css/mdb.min.css';
import { MDBDataTable } from 'mdbreact';
import Container from 'react-bootstrap/Container';
import Select from "react-select";

function App() {
  document.title = "FTI Amazonsales Record";

  const [data, setData] = useState([]);
  const labelIndex = 8;
  const countryOption = [{ value: 'JP', label: 'JP' }, { value: 'US', label: 'US' }];
  const productInfo = {
    'JP': {
      '98100C220Z0000': { 'id': '98100C220Z0000', 'name': 'INNEX C220' },
      '98100C830Z0000': { 'id': '98100C830Z0000', 'name': 'INNEX C830' },
      '98300DC4000100': { 'id': '98300DC4000100', 'name': 'IDEAO DC400 (PAL)' },
      '98100C470Z0100': { 'id': '98100C470Z0100', 'name': 'INNEX C470 (PAL)' },
      '98100C840Z0000': { 'id': '98100C840Z0000', 'name': 'INNEX Cube' },
      '98200PM2400001': { 'id': '98200PM2400001', 'name': 'IDEAO Hub with Stand JP (PM-240)' },
      '98900P001Z0000': { 'id': '98900P001Z0000', 'name': 'IDEAO Pen' },
      '32-WXCG-5EHJ': { 'id': '98100C570Z0000', 'name': 'INNEX C570' },
      '98300DC5000002': { 'id': '98300DC5000000', 'name': 'INNEX DC500' },
    }, 'US': {
      '98100C220Z0000': { 'id': '98100C220Z0000', 'name': 'INNEX C220' },
      '98100C830Z0001': { 'id': '98100C830Z0000', 'name': 'INNEX C830' },
      '98100C470Z0000': { 'id': '98100C470Z0000', 'name': 'INNEX C470 (USA)' },
      '98300DC4000000': { 'id': '98300DC4000000', 'name': 'IDEAO DC400 (USA)' },
      '98100C840Z0000': { 'id': '98100C840Z0000', 'name': 'INNEX Cube' },
      '98200PM2400000': { 'id': '98200PM2400000', 'name': 'IDEAO Hub with Stand (PM-240)' },
      '98900P001Z0000': { 'id': '98900P001Z0000', 'name': 'IDEAO Pen' },
      '98100C570Z0000': { 'id': '98100C570Z0000', 'name': 'INNEX C570' },
      '98300DC5000002': { 'id': '98300DC5000000', 'name': 'INNEX DC500' },
    }
  };
  const months = {
    Jan: '01',
    Feb: '02',
    Mar: '03',
    Apr: '04',
    May: '05',
    Jun: '06',
    Jul: '07',
    Aug: '08',
    Sep: '09',
    Oct: '10',
    Nov: '11',
    Dec: '12'
  };

  const handleFileUpload = (e) => {
    const reader = new FileReader();
    reader.readAsBinaryString(e.target.files[0]);
    reader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const dataArr = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      const dataHeader = dataArr[labelIndex - 1];
      const parsedData = XLSX.utils.sheet_to_json(sheet, { header: dataHeader });

      let col = Object.keys(parsedData[labelIndex]).map(function (val, i) {
        return { 'label': val, 'field': val, 'width': 200 };
      })

      const mdbData = { columns: col, rows: parsedData.slice(labelIndex, parsedData.length) };
      const dataSet = { 'mdbData': mdbData, 'dataArr': dataArr };
      setData(dataSet);

      document.getElementById("submitBtn").style.display = "block";
    };
  }

  const [selected, setSelected] = useState(null);
  const handleSelect = (selectedOption) => {
    setSelected(selectedOption);
  };

  function handleSubmit() {
    const date_ = new Date().toLocaleDateString().toString();
    const date = date_.split('/')[2] + date_.split('/')[0].padStart(2, '0') + date_.split('/')[1];

    const dataArr = data['dataArr'].slice(labelIndex, data['dataArr'].length);
    var wsData = { 'bill': [], 'certificate': [] };

    const country = selected['value'];

    function getI(char) {
      return char.charCodeAt(0) - ('A').charCodeAt(0);
    }

    if (country === "JP") {
      wsData['bill'].push(['日期', '序號', '客戶/供應商編碼', '客戶/供應商名稱', '承辦人', '發貨倉庫', '交易類型', '貨幣', '匯率', '銷貨單號碼', '品項編碼', '品項名', '規格', '數量', '單價', '外幣金額', '稅前價格', '營業稅', '摘要', '產生生產入庫', '折扣']);
      wsData['certificate'].push(['憑證日期', '序號', '會計憑證號碼', '營業稅類型', '客戶/供應商編碼', '客戶/供應商名稱', '稅前價格', '外幣金額', '匯率', '營業稅', '類型', '科目編碼', '客戶/供應商編碼', '客戶/供應商名稱', '借方', '貸方', '外幣金額', '匯率', '摘要編碼', '摘要'])
      let cnt = 1;
      let lastDate = "";
      let dateCnt = 0;
      dataArr.forEach(function (row, array) {
        if (row[getI('C')] != '注文') return;
        console.log(row);
        const yymmdd_ = row[getI('A')].split(" ")[0].split("/");
        const yymmdd = yymmdd_[0] + yymmdd_[1].padStart(2, '0') + yymmdd_[2].padStart(2, '0');
        // const yymmddLs = row[getI('A')].split(" ")[0].split("/");

        //bill
        wsData['bill'].push([
          yymmdd,
          cnt.toString().padStart(3, '0'),
          'SUP00038',
          'Amazon.co.jp',
          'FTI10011',
          '107',
          '12',
          '00004',
          '',
          row[getI('D')],
          productInfo[country][row[getI('E')]]['id'],
          productInfo[country][row[getI('E')]]['name'],
          '',
          row[getI('G')],
          (parseFloat(row[getI('N')]) / parseFloat(row[getI('G')])).toFixed(2),
          { t: "n", f: "ROUND(N" + (cnt + 1).toString() + "*O" + (cnt + 1).toString() + ",2)" },
          { t: "n", f: "ROUND(I" + (cnt + 1).toString() + "*P" + (cnt + 1).toString() + ",1)" },
          '',
          '',
          '',
          row[getI('U')]
        ]);

        //certificate
        const foreignCurrency = [
          parseFloat(row[27]) - parseFloat(row[getI('O')]) - parseFloat(row[getI('V')]),
          (parseFloat(row[getI('T')]) + parseFloat(row[getI('X')])) * (-1),
          (parseFloat(row[getI('Y')])) * (-1)
        ];
        let certificateNo = '';
        if (yymmdd == lastDate) {
          dateCnt += 1;
          certificateNo = 'Amazon' + yymmdd + '30' + dateCnt.toString().padStart(2, '0');
        }
        else {
          lastDate = yymmdd;
          dateCnt = 1;
          certificateNo = 'Amazon' + yymmdd + '30' + dateCnt.toString().padStart(2, '0');
        }
        wsData['certificate'].push([
          yymmdd,
          cnt.toString().padStart(3, '0'),
          certificateNo,
          'Y1',
          'SUP00038',
          'Amazon.co.jp',
          { t: "n", f: "ROUND(H" + ((cnt - 1) * 4 + 1 + 1).toString() + "*I" + ((cnt - 1) * 4 + 1 + 1).toString() + ",1)" },
          (parseFloat(row[getI('N')]) + parseFloat(row[getI('U')])).toFixed(2),
          '',
          '',
          '3', //diff. from this line
          '1141',
          'SUP00038',
          'Amazon.co.jp',
          { t: "n", f: "ROUND(Q" + ((cnt - 1) * 4 + 1 + 1).toString() + "*R" + ((cnt - 1) * 4 + 1 + 1).toString() + ",1)" },
          '',
          foreignCurrency[0].toFixed(2),
          '',
          '',
          row[getI('D')] + " " + productInfo[country][row[getI('E')]].name
        ]);
        wsData['certificate'].push([yymmdd, cnt.toString().padStart(3, '0'), certificateNo, 'Y1', 'SUP00038', 'Amazon.co.jp', { t: "n", f: "ROUND(H" + ((cnt - 1) * 4 + 1 + 2).toString() + "*I" + ((cnt - 1) * 4 + 1 + 2).toString() + ",1)" }, (parseFloat(row[getI('N')]) + parseFloat(row[getI('U')])).toFixed(2), '', '',
          '3', //diff. from this line
          '611801',
          '42527250',
          'Amazon Japan GK',
          { t: "n", f: "ROUND(Q" + ((cnt - 1) * 4 + 1 + 2).toString() + "*R" + ((cnt - 1) * 4 + 1 + 2).toString() + ",1)" },
          '',
          foreignCurrency[1].toFixed(2),
          '', '', row[getI('D')] + " " + productInfo[country][row[getI('E')]].name
        ]);
        wsData['certificate'].push([yymmdd, cnt.toString().padStart(3, '0'), certificateNo, 'Y1', 'SUP00038', 'Amazon.co.jp', { t: "n", f: "ROUND(H" + ((cnt - 1) * 4 + 1 + 3).toString() + "*I" + ((cnt - 1) * 4 + 1 + 3).toString() + ",1)" }, (parseFloat(row[getI('N')]) + parseFloat(row[getI('U')])).toFixed(2), '', '',
          '3', //diff. from this line
          '6115',
          '42527250',
          'Amazon Japan GK',
          { t: "n", f: "ROUND(Q" + ((cnt - 1) * 4 + 1 + 3).toString() + "*R" + ((cnt - 1) * 4 + 1 + 3).toString() + ",1)" },
          '',
          foreignCurrency[2].toFixed(2),
          '', '', row[getI('D')] + " " + productInfo[country][row[getI('E')]].name
        ]);
        wsData['certificate'].push([yymmdd, cnt.toString().padStart(3, '0'), certificateNo, 'Y1', 'SUP00038', 'Amazon.co.jp', { t: "n", f: "ROUND(H" + ((cnt - 1) * 4 + 1 + 4).toString() + "*I" + ((cnt - 1) * 4 + 1 + 4).toString() + ",1)" }, (parseFloat(row[getI('N')]) + parseFloat(row[getI('U')])).toFixed(2), '', '',
          '4', //diff. from this line
          '4111',
          'SUP00038',
          'Amazon.co.jp',
          '',
          { t: "n", f: "ROUND(Q" + ((cnt - 1) * 4 + 1 + 4).toString() + "*R" + ((cnt - 1) * 4 + 1 + 4).toString() + ",1)" },
          (foreignCurrency[0] + foreignCurrency[1] + foreignCurrency[2]).toFixed(2),
          '', '', row[getI('D')] + " " + productInfo[country][row[getI('E')]].name
        ]);
        cnt += 1;
      });
    }
    else if (country === 'US') {
      wsData['bill'].push(['日期', '序號', '客戶/供應商編碼', '客戶/供應商名稱', '承辦人', '發貨倉庫', '交易類型', '貨幣', '匯率', '銷貨單號碼', '品項編碼', '品項名', '規格', '數量', '單價', '外幣金額', '稅前價格', '營業稅', '摘要', '產生生產入庫']);
      wsData['certificate'].push(['憑證日期', '序號', '會計憑證號碼', '營業稅類型', '客戶/供應商編碼', '客戶/供應商名稱', '稅前價格', '外幣金額', '匯率', '營業稅', '類型', '科目編碼', '客戶/供應商編碼', '客戶/供應商名稱', '借方', '貸方', '外幣金額', '匯率', '摘要編碼', '摘要'])
      let cnt = 1;
      let lastDate = "";
      let dateCnt = 0;
      dataArr.forEach(function (row, array) {
        if (row[getI('C')] != 'Order') return;
        const sub = row[getI('A')].split(" ")
        const yymmdd = sub[2] + (months[sub[0]]).padStart(2, '0') + (sub[1].split(",")[0]).padStart(2, '0');

        //bill
        wsData['bill'].push([
          yymmdd,
          cnt.toString().padStart(3, '0'),
          'SUP00022',
          'Amazon.com',
          'FTI10011',
          '102',
          '12',
          '00001',
          '',
          row[getI('D')],
          productInfo[country][row[getI('E')]]['id'],
          productInfo[country][row[getI('E')]]['name'],
          '',
          row[getI('G')],
          (parseFloat(row[getI('O')]) / parseFloat(row[getI('G')])).toFixed(2),
          { t: "n", f: "ROUND(N" + (cnt + 1).toString() + "*O" + (cnt + 1).toString() + ",2)" },
          { t: "n", f: "ROUND(I" + (cnt + 1).toString() + "*P" + (cnt + 1).toString() + ",1)" },
          '',
          '',
          ''
        ]);

        //certificate
        const foreignCurrency = [
          parseFloat(row[29]),
          (parseFloat(row[getI('Z')])) * (-1),
          (parseFloat(row[26])) * (-1)
        ];
        let certificateNo = '';
        if (yymmdd == lastDate) {
          dateCnt += 1;
          certificateNo = 'Amazon' + yymmdd + '30' + dateCnt.toString().padStart(2, '0');
        }
        else {
          lastDate = yymmdd;
          dateCnt = 1;
          certificateNo = 'Amazon' + yymmdd + '30' + dateCnt.toString().padStart(2, '0');
        }
        wsData['certificate'].push([
          yymmdd,
          cnt.toString().padStart(3, '0'),
          certificateNo,
          'Y1',
          'SUP00022',
          'Amazon.com',
          { t: "n", f: "ROUND(H" + ((cnt - 1) * 4 + 1 + 1).toString() + "*I" + ((cnt - 1) * 4 + 1 + 1).toString() + ",1)" },
          row[getI('O')],
          '',
          '',
          '3', //diff. from this line
          '1141',
          'SUP00022',
          'Amazon.com',
          { t: "n", f: "ROUND(Q" + ((cnt - 1) * 4 + 1 + 1).toString() + "*R" + ((cnt - 1) * 4 + 1 + 1).toString() + ",1)" },
          '',
          foreignCurrency[0].toFixed(2),
          '',
          '',
          row[getI('D')] + " " + productInfo[country][row[getI('E')]].name
        ]);
        wsData['certificate'].push([yymmdd, cnt.toString().padStart(3, '0'), certificateNo, 'Y1', 'SUP00022', 'Amazon.com', { t: "n", f: "ROUND(H" + ((cnt - 1) * 4 + 1 + 2).toString() + "*I" + ((cnt - 1) * 4 + 1 + 2).toString() + ",1)" }, row[getI('O')], '', '',
          '3', //diff. from this line
          '611801',
          '76315326',
          'Amazon.com Services LLC',
          { t: "n", f: "ROUND(Q" + ((cnt - 1) * 4 + 1 + 2).toString() + "*R" + ((cnt - 1) * 4 + 1 + 2).toString() + ",1)" },
          '',
          foreignCurrency[1].toFixed(2),
          '', '', row[getI('D')] + " " + productInfo[country][row[getI('E')]].name
        ]);
        wsData['certificate'].push([yymmdd, cnt.toString().padStart(3, '0'), certificateNo, 'Y1', 'SUP00022', 'Amazon.com', { t: "n", f: "ROUND(H" + ((cnt - 1) * 4 + 1 + 3).toString() + "*I" + ((cnt - 1) * 4 + 1 + 3).toString() + ",1)" }, row[getI('O')], '', '',
          '3', //diff. from this line
          '6115',
          '76315326',
          'Amazon.com Services LLC',
          { t: "n", f: "ROUND(Q" + ((cnt - 1) * 4 + 1 + 3).toString() + "*R" + ((cnt - 1) * 4 + 1 + 3).toString() + ",1)" },
          '',
          foreignCurrency[2].toFixed(2),
          '', '', row[getI('D')] + " " + productInfo[country][row[getI('E')]].name
        ]);
        wsData['certificate'].push([yymmdd, cnt.toString().padStart(3, '0'), certificateNo, 'Y1', 'SUP00022', 'Amazon.com', { t: "n", f: "ROUND(H" + ((cnt - 1) * 4 + 1 + 4).toString() + "*I" + ((cnt - 1) * 4 + 1 + 4).toString() + ",1)" }, row[getI('O')], '', '',
          '4', //diff. from this line
          '4111',
          'SUP00022',
          'Amazon.com',
          '',
          { t: "n", f: "ROUND(Q" + ((cnt - 1) * 4 + 1 + 4).toString() + "*R" + ((cnt - 1) * 4 + 1 + 4).toString() + ",1)" },
          (foreignCurrency[0] + foreignCurrency[1] + foreignCurrency[2]).toFixed(2),
          '', '', row[getI('D')] + " " + productInfo[country][row[getI('E')]].name
        ]);
        cnt += 1;
      });
    }
    else {
      alert('selected country invalid');
      return;
    }

    var wb = XLSX.utils.book_new(),
      wsBill = XLSX.utils.aoa_to_sheet(wsData['bill']),
      wsCertificate = XLSX.utils.aoa_to_sheet(wsData['certificate']);
    Object.keys(wsBill).forEach(key => {
      if (!key.startsWith("!")) {
        wsBill[key].s = {
          alignment: {
            vertical: "center"
          }
        };
      }
    })
    Object.keys(wsCertificate).forEach(key => {
      if (!key.startsWith("!")) {
        wsCertificate[key].s = {
          alignment: {
            vertical: "center"
          }
        };
      }
    })
    XLSX.utils.book_append_sheet(wb, wsBill, "銷貨單");
    XLSX.utils.book_append_sheet(wb, wsCertificate, "應收憑證 II");
    XLSX.writeFile(wb, "Amazon " + country + " " + date + ".xlsx");
  }


  const customStyles = {
    option: (styles, state) => ({
      ...styles,
      color: state.isSelected ? "#FFF" : styles.color,
      backgroundColor: state.isSelected ? "#767171" : styles.color,
      borderBottom: "1px solid rgba(0, 0, 0, 0.125)",
      "&:hover": {
        color: "#FFF",
        backgroundColor: "#767171"
      }
    }),
    control: (styles, state) => ({
      ...styles,
      boxShadow: state.isFocused ? "0 0 0 0.2rem rgba(255, 131, 41, 0.25)" : 0,
      borderColor: state.isFocused ? "#FF8329" : "#CED4DA",
      "&:hover": {
        borderColor: state.isFocused ? "#FF8329" : "#CED4DA"
      }
    })
  };

  return (
    <Container>
      <center>
        <br></br><br></br>
        <h1>Custom Unified Transaction File Upload</h1>
        <br></br>
        <Select
          styles={customStyles}
          options={countryOption}
          onChange={handleSelect}
          autoFocus={true}
        />
        <br></br>
        <input
          type="file"
          accept=".xlsx, .xls, .csv"
          onChange={handleFileUpload}
        />
        <br></br><br></br>
        <button style={{ "display": "none" }} onClick={handleSubmit} id="submitBtn">download</button>
        <br></br>
      </center>

      {
        Object.keys(data).length > 0 && (
          <MDBDataTable
            scrollX
            striped
            bordered
            data={data['mdbData']}
            entriesOptions={[1, 2, 5, 10, 25, 100]}
            entries={5}
          />
        )
      }

    </Container >
  );
}


export default App;

