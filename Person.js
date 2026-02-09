const mongoose = require('mongoose');

const personSchema = new mongoose.Schema({
  idNumber: { type: String, required: true, unique: true }, // 身分證
  name: String,
  birth: String,         // 生日可用 String 儲存 YYYY-MM-DD
  education: String,     // 學歷
  phone: String,
  address: String,
  dateUpdated: String,   // 更新日期 YYYY-MM-DD
  items: {
    healthCheck: { type: Boolean, default: false },
    bc: { type: Boolean, default: false },
    papSmear: { type: Boolean, default: false },
    hpv: { type: Boolean, default: false },
    colonScreen: { type: Boolean, default: false },
    oralScreen: { type: Boolean, default: false },
    icp: { type: Boolean, default: false },
    gastricCancer: { type: Boolean, default: false }
  }
});

module.exports = mongoose.model('Person', personSchema);
