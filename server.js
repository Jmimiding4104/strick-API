const express = require('express');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const Person = require('./Person');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// 中介軟體
app.use(bodyParser.json());

// 連接 MongoDB
mongoose.connect('mongodb://127.0.0.1:27017/personsDB', {
  useNewUrlParser: true,
  useUnifiedTopology: true
});

mongoose.connection.on('connected', () => {
    console.log('✅ MongoDB 連線成功');
  });
  
  mongoose.connection.on('error', (err) => {
    console.error('❌ MongoDB 連線錯誤:', err);
  });

// 查詢個人資料
app.get('/person/:idNumber', async (req, res) => {
  const person = await Person.findOne({ idNumber: req.params.idNumber });
  if (person) {
    const { name, birth, education, phone, address } = person;
    res.json({ name, birth, education, phone, address, items: person.items });
  } else {
    res.status(404).json({ message: '查無資料' });
  }
});

// 新增或更新個人資料
app.post('/person', async (req, res) => {
  const {
    idNumber, name, birth, education, phone, address,
    items = {} // healthCheck, bc, papSmear, hpv, colonScreen, oralScreen, icp, gastricCancer
  } = req.body;

  const now = new Date();
  now.setHours(now.getHours() + 8); // 將時間調整為 UTC+8
  const today = now.toISOString().split('T')[0];

  let person = await Person.findOne({ idNumber });

  if (person) {
    // 更新
    person.name = name;
    person.birth = birth;
    person.education = education;
    person.phone = phone;
    person.address = address;
    person.dateUpdated = today;
    person.items = { ...person.items, ...items };
    await person.save();
    res.json({ message: '更新成功' });
  } else {
    // 新增
    person = new Person({
      idNumber, name, birth, education, phone, address,
      dateUpdated: today,
      items
    });
    await person.save();
    res.json({ message: '新增成功' });
  }
});

app.get('/export', async (req, res) => {
    try {
      const { date } = req.query; // 從 query 取得日期
      let filter = {};
  
      if (date) {
        filter.dateUpdated = date; // 篩選特定日期
      }
  
      const persons = await Person.find(filter);
  
      if (persons.length === 0) {
        return res.status(404).send('查無資料');
      }
  
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('People');
  
      // 設定標題列
      worksheet.columns = [
        { header: '身分證字號', key: 'idNumber', width: 15 },
        { header: '姓名', key: 'name', width: 10 },
        { header: '生日(年)', key: 'birthYear', width: 10 },
        { header: '生日(月)', key: 'birthMonth', width: 10 },
        { header: '生日(日)', key: 'birthDay', width: 10 },
        { header: '學歷', key: 'education', width: 10 },
        { header: '電話', key: 'phone', width: 15 },
        { header: '住址', key: 'address', width: 30 },
        { header: '更新日期', key: 'dateUpdated', width: 12 },
        { header: '健檢', key: 'healthCheck', width: 8 },
        { header: 'BC', key: 'bc', width: 8 },
        { header: '子抹', key: 'papSmear', width: 8 },
        { header: 'HPV', key: 'hpv', width: 8 },
        { header: '腸篩', key: 'colonScreen', width: 8 },
        { header: '口篩', key: 'oralScreen', width: 8 },
        { header: 'ICP', key: 'icp', width: 8 },
        { header: '胃癌', key: 'gastricCancer', width: 8 }
      ];
  
      // 寫入資料列
      persons.forEach(person => {

        let birthYear = '', birthMonth = '', birthDay = '';
        if (person.birth && person.birth.length === 6) {
          birthYear = person.birth.slice(0, 2);
          birthMonth = person.birth.slice(2, 4);
          birthDay = person.birth.slice(4, 6);
        }

        worksheet.addRow({
          idNumber: person.idNumber,
          name: person.name,
          birthYear,
          birthMonth,
          birthDay,
          birth: person.birth,
          education: person.education,
          phone: person.phone,
          address: person.address,
          dateUpdated: person.dateUpdated,
          healthCheck: person.items?.healthCheck ?? false,
          bc: person.items?.bc ?? false,
          papSmear: person.items?.papSmear ?? false,
          hpv: person.items?.hpv ?? false,
          colonScreen: person.items?.colonScreen ?? false,
          oralScreen: person.items?.oralScreen ?? false,
          icp: person.items?.icp ?? false,
          gastricCancer: person.items?.gastricCancer ?? false
        });
      });
  
      const filename = `export_${date || 'all'}.xlsx`;
      const exportPath = path.join(process.cwd(), 'export.xlsx');
      await workbook.xlsx.writeFile(exportPath);
  
      res.download(exportPath, `人員資料匯出_${date || '全部'}.xlsx`, (err) => {
        if (err) {
          console.error('下載時發生錯誤:', err);
          res.status(500).send('下載失敗');
        } else {
          fs.unlinkSync(exportPath); // 匯出完即刪除
        }
      });
    } catch (error) {
      console.error('匯出失敗:', error);
      res.status(500).send('伺服器錯誤');
    }
  });
  

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
