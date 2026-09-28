/** GD Record | Bound to สำเนาของ Record Data GD. */

/**

 * ระบบบันทึกการปฏิบัติงานบน Google Apps Script

 *

 * ภาพรวมการทำงาน:

 * 1. onOpen / showGDRecord / doGet เปิดหน้าจอจากไฟล์ HTML ชื่อ Index

 * 2. getBootstrap อ่านข้อมูลทุกตารางเพื่อส่งให้หน้าจอ

 * 3. saveRecord ตรวจสอบและบันทึกลูกค้า ผู้ติดต่อ หรือรายงานการปฏิบัติงาน

 * 4. auditData ตรวจสอบความสัมพันธ์ของข้อมูลโดยไม่เพิ่มข้อมูลทดสอบ

 *

 * ชื่อฟังก์ชันที่ลงท้ายด้วย _ เป็นฟังก์ชันช่วยภายใน

 * ฉบับนี้จัดรูปแบบและเพิ่มคำอธิบาย โดยคงคำสั่งและเงื่อนไขเดิมไว้

 */


// ============================================================================

// 1. การตั้งค่าและโครงสร้างตาราง

// ============================================================================


// id คือรหัส Spreadsheet ที่ใช้เก็บข้อมูล ส่วน tz คือเขตเวลาเริ่มต้น

// Object.freeze ป้องกันการแก้ไข property ระดับแรกของออบเจ็กต์

const GD = Object.freeze({

  id: '157Ikgt7YALWWoRvSnaNvv0uaMICPgvZk-XKT2-tMjKk',

  tz: 'Asia/Bangkok'

});


// ชื่อ property ต้องตรงกับชื่อชีต และหัวคอลัมน์ต้องตรงตามลำดับนี้

// คอลัมน์แรกของแต่ละตารางใช้ระบุรายการและหาแถวสุดท้ายที่มีข้อมูล

// visit_at_th เป็นคอลัมน์สูตร จึงไม่เขียนทับเมื่อเพิ่มรายงาน

// หมายเหตุ: Object.freeze ชั้นนอกไม่ได้ freeze อาร์เรย์หัวคอลัมน์ภายใน

const HEADERS = Object.freeze({

  customer: [

    'customer_id',

    'company_name',

    'company_short_name',

    'customer_type',

    'phone',

    'address',

    'province',

    'location_url',

    'status',

    'remark'

  ],

  customer_contact: [

    'contact_id',

    'customer_id',

    'ชื่อ',

    'นามสกุล',

    'ตำแหน่ง',

    'โทรศัพท์',

    'email'

  ],

  reports: [

    'report_id',

    'visit_at',

    'customer_id',

    'contact_id',

    'activity',

    'job_detail',

    'remark',

    'status',

    'next_follow_up',

    'attachment_url',

    'created_by',

    'created_at',

    'updated_at',

    'visit_at_th'

  ],

  employee: [

    'employee_id',

    'ชื่อ',

    'นามสกุล',

    'ชื่อย่อ',

    'ตำแหน่ง',

    'status'

  ],

  report_employee: [

    'report_id',

    'employee_id',

    'report_employee_id'

  ]

});


// ============================================================================

// 2. การเปิดหน้าจอ

// ============================================================================


/** เพิ่มเมนู GD Record เมื่อเปิด Spreadsheet ที่ผูกกับสคริปต์ */

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu('GD Record')

    .addItem('เปิดหน้าบันทึกข้อมูล', 'showGDRecord')

    .addToUi();

}


/** เปิด Index.html เป็นหน้าต่างใน Spreadsheet ขนาด 1180 × 820 */

function showGDRecord() {

  SpreadsheetApp.getUi().showModalDialog(

    HtmlService.createHtmlOutputFromFile('Index')

      .setWidth(1180)

      .setHeight(820),

    'GD Record · บันทึกการปฏิบัติงาน'

  );

}


/** จุดรับคำขอ GET เมื่อเผยแพร่สคริปต์เป็น Web App */

function doGet() {

  return HtmlService.createHtmlOutputFromFile('Index')

    .setTitle('GD Record')

    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

}


// ============================================================================

// 3. การอ่านข้อมูลและเตรียมข้อมูลให้หน้าจอ

// ============================================================================


/** เปิดฐานข้อมูลตามรหัสที่กำหนดไว้ใน GD.id */

function db_() {

  return SpreadsheetApp.openById(GD.id);

}


/**

 * อ่านชีตและตรวจหัวคอลัมน์ก่อนใช้งาน

 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} db Spreadsheet ที่เปิดแล้ว

 * @param {string} name ชื่อตารางที่ประกาศใน HEADERS

 * @return {Object} sheet, headers, last และ rows

 * last คือเลขแถวสุดท้ายที่พบค่าจากคอลัมน์แรก (อย่างน้อย 1 คือแถวหัวตาราง)

 * rows คืออาร์เรย์ออบเจ็กต์ เช่น { customer_id: 'CUS_...', company_name: '...' }

 */

function table_(db, name) {

  const sheet = db.getSheetByName(name);

  if (!sheet) throw Error('ไม่พบตาราง ' + name);


  const headers = sheet

    .getRange(1, 1, 1, HEADERS[name].length)

    .getDisplayValues()[0];


  if (headers.some((v, i) => v !== HEADERS[name][i]))

    throw Error('โครงสร้างตาราง ' + name + ' เปลี่ยนไป กรุณาตรวจสอบหัวคอลัมน์');


  // Do not use the formula-spill column to determine the append row.

  // สูตรอาจแสดงผลลงไปไกลกว่าแถวข้อมูลจริง จึงใช้คอลัมน์รหัสหาแถวสำหรับเพิ่มข้อมูล

  const ids = sheet

    .getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1)

    .getDisplayValues();


  let last = ids.length;

  while (last > 1 && !ids[last - 1][0]) last--;


  const rows = last > 1

    ? sheet.getRange(2, 1, last - 1, headers.length).getValues()

    : [];


  // ข้ามแถวที่รหัสว่างหลังตัดช่องว่าง และจับคู่ชื่อคอลัมน์กับค่าของแต่ละแถว

  return {

    sheet,

    headers,

    last,

    rows: rows

      .filter(r => String(r[0]).trim())

      .map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])))

  };

}


/** สถานะว่างหรือ ACTIVE ถือว่าใช้งานได้ โดย ACTIVE ไม่แยกตัวพิมพ์ใหญ่/เล็ก */

function active_(r) {

  return !r.status || String(r.status).toUpperCase() === 'ACTIVE';

}


/**

 * แปลงค่าจากชีตเป็นข้อความก่อนส่งให้หน้าจอ

 * Date ใช้รูปแบบ yyyy-MM-dd HH:mm; null/undefined กลายเป็นข้อความว่าง

 * ตามเงื่อนไขเดิม หากปีที่จัดรูปแบบได้ตั้งแต่ 2400 ขึ้นไป จะลบ 543

 */

function display_(v, tz) {

  if (v instanceof Date) {

    // The spreadsheet locale can expose no timezone to a web-app execution.

    // Always pass a concrete IANA timezone to Utilities.formatDate.

    // หากไม่ได้รับเขตเวลาเป็นข้อความที่ไม่ว่าง ให้ใช้ Asia/Bangkok จาก GD.tz

    const zone = (typeof tz === 'string' && tz) ? tz : GD.tz;

    let s = Utilities.formatDate(v, zone, 'yyyy-MM-dd HH:mm');

    const year = +s.slice(0, 4);


    if (year >= 2400) s = String(year - 543) + s.slice(4);

    return s;

  }


  return v == null ? '' : String(v);

}


/**

 * โหลดทุกตารางใน HEADERS และแปลงค่าทุกช่องเป็นข้อความ

 * เพิ่ม activities จากรายการพื้นฐานรวมกับกิจกรรมในรายงานเดิม โดยตัดค่าซ้ำ

 * ฟังก์ชันนี้คืนข้อมูลทั้งหมด ไม่ได้กรองเฉพาะรายการ ACTIVE

 */

function getBootstrap() {

  const db = db_(), data = {}, tz = db.getSpreadsheetTimeZone();


  Object.keys(HEADERS).forEach(n => {

    data[n] = table_(db, n).rows.map(r =>

      Object.fromEntries(

        Object.entries(r).map(([k, v]) => [k, display_(v, tz)])

      )

    );

  });


  data.activities = [...new Set([

    'Site Survey',

    'Installation',

    'Visit',

    'Maintenance',

    'Check Communication',

    'Meeting',

    'Training',

    'Follow Up',

    'Commissioning',

    'Troubleshooting',

    ...data.reports.map(r => r.activity).filter(Boolean)

  ])];


  return data;

}


// ============================================================================

// 4. ตัวช่วยตรวจสอบและแปลงค่าก่อนบันทึก

// ============================================================================


/**

 * อ่าน p[k] เป็นข้อความและตัดช่องว่างหัวท้าย

 * required กำหนดว่าห้ามว่าง; max คือความยาวสูงสุด (ค่าเริ่มต้น 2000)

 * การนับความยาวใช้ String.length ตามพฤติกรรม JavaScript เดิม

 */

function text_(p, k, required, max) {

  const v = String(p[k] == null ? '' : p[k]).trim();

  if (required && !v) throw Error('กรุณากรอก ' + k);

  if (v.length > (max || 2000)) throw Error(k + ' ยาวเกินกำหนด');

  return v;

}


/** ค่าต้องไม่ว่างและต้องตรงกับสมาชิกใน allowed แบบตรงตัว */

function option_(p, k, allowed) {

  const v = text_(p, k, true);

  if (!allowed.includes(v)) throw Error('ค่า ' + k + ' ไม่ถูกต้อง');

  return v;

}


/** ลิงก์เว้นว่างได้; หากกรอกต้องขึ้นต้น http:// หรือ https:// และไม่มีช่องว่าง */

function url_(p, k) {

  const v = text_(p, k, false, 2000);

  if (v && !/^https?:\/\/[^\s]+$/i.test(v))

    throw Error('ลิงก์ต้องขึ้นต้นด้วย https:// หรือ http://');

  return v;

}


/**

 * ค้นหารายการอ้างอิงจาก rows โดยใช้ key และ id

 * label ใช้ประกอบข้อความแจ้งข้อผิดพลาด

 * onlyActive เป็นจริงเมื่อต้องตรวจสถานะด้วย active_

 */

function ref_(rows, key, id, label, onlyActive) {

  const r = rows.find(x => String(x[key]) === id);

  if (!r || (onlyActive && !active_(r)))

    throw Error('ไม่พบ' + label + 'ที่ใช้งานได้ กรุณาโหลดข้อมูลใหม่');

  return r;

}


/**

 * แปลงวันที่เป็นเลขวันที่ของ Google Sheets

 * withTime=true รับ YYYY-MM-DDTHH:mm; false รับ YYYY-MM-DD

 * อนุญาตปี ค.ศ. 2000–2200 และตรวจวันจริง เช่น ปฏิเสธ 31 กุมภาพันธ์

 * Date.UTC ใช้คำนวณจากองค์ประกอบวันที่โดยไม่ผูกกับเขตเวลาของเครื่อง

 * 86400000 คือมิลลิวินาทีต่อวัน; 25569 คือค่า serial ของวันที่ 1970-01-01

 * เศษทศนิยมของผลลัพธ์คือเวลาในวันนั้น ไม่ใช่การแปลงเวลาไทยให้เป็น UTC

 */

function serial_(value, withTime) {

  const re = withTime

    ? /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/

    : /^(\d{4})-(\d{2})-(\d{2})$/;


  const m = String(value).match(re);

  if (!m) throw Error('รูปแบบวันที่ไม่ถูกต้อง');


  const [y, mo, d, h, mi] = [

    +m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0)

  ];

  const dt = new Date(Date.UTC(y, mo - 1, d, h, mi));


  if (

    y < 2000 ||

    y > 2200 ||

    dt.getUTCFullYear() !== y ||

    dt.getUTCMonth() !== mo - 1 ||

    dt.getUTCDate() !== d ||

    h > 23 ||

    mi > 59

  ) throw Error('วันที่ไม่ถูกต้อง กรุณาใช้ปี ค.ศ.');


  return dt.getTime() / 86400000 + 25569;

}


/** เติม ' หน้าข้อความที่ขึ้นต้น = + @ - เพื่อให้ชีตรับเป็นข้อความแทนสูตร */

function safe_(v) {

  return typeof v === 'string' && /^[=+@-]/.test(v) ? "'" + v : v;

}


/**

 * เพิ่มข้อมูลหนึ่งแถวและเก็บช่วงเซลล์ไว้ใน undo สำหรับล้างเมื่อเกิดข้อผิดพลาด

 * t คือผลจาก table_; obj คือข้อมูลตามชื่อคอลัมน์; คืนค่าเลขแถวที่เขียน

 * ไม่เขียน visit_at_th โดยอาศัยโครงสร้างเดิมที่วางคอลัมน์นี้ไว้ท้ายสุด

 */

function append_(t, obj, undo) {

  const row = t.last + 1,

    width = t.headers.filter(h => h !== 'visit_at_th').length;


  // หากแถวไม่พอ ให้เพิ่มอย่างน้อย 100 แถว หรือเท่าที่จำเป็นหากมากกว่า 100

  if (row > t.sheet.getMaxRows())

    t.sheet.insertRowsAfter(

      t.sheet.getMaxRows(),

      Math.max(100, row - t.sheet.getMaxRows())

    );


  const range = t.sheet.getRange(row, 1, 1, width);

  undo.push(range);


  // ค่าที่ไม่ได้ระบุใช้ข้อความว่าง และจัดลำดับค่าตามหัวคอลัมน์เสมอ

  range.setValues([

    t.headers.slice(0, width).map(h => safe_(obj[h] === undefined ? '' : obj[h]))

  ]);

  t.last = row;


  return row;

}


// ============================================================================

// 5. บันทึกข้อมูลหลัก

// ============================================================================


/** Token-derived IDs make retries safe even after the response is lost. */

/**

 * บันทึกข้อมูลที่หน้าจอส่งมา

 * @param {string} kind customer, customer_contact หรือ reports

 * @param {Object} p ข้อมูลฟอร์ม (payload)

 * @param {string} token รหัสคำขอเลขฐานสิบหก 32 ตัวอักษร

 * @return {Object} { id, duplicate } โดย duplicate=true หมายถึงพบรหัสเดิมแล้ว

 *

 * ใช้ token สร้างรหัสรายการ: การส่งซ้ำด้วย kind/token เดิมจะไม่เพิ่มรายการใหม่

 * การตรวจซ้ำดูเฉพาะรหัส ไม่ได้เปรียบเทียบ payload หรือซ่อมข้อมูลที่บันทึกไม่ครบ

 * ScriptLock ช่วยกันคำขอในสคริปต์เดียวกันเขียนพร้อมกันระหว่างถือ lock

 * undo ช่วยล้างเนื้อหาที่เพิ่งเขียนหากผิดพลาด แต่ไม่ใช่ธุรกรรมฐานข้อมูลเต็มรูปแบบ

 */

function saveRecord(kind, p, token) {

  if (!['customer', 'customer_contact', 'reports'].includes(kind))

    throw Error('ประเภทข้อมูลไม่ถูกต้อง');


  if (!/^[a-f0-9]{32}$/i.test(token))

    throw Error('รหัสคำขอไม่ถูกต้อง กรุณาเปิดฟอร์มใหม่');


  p = p || {};

  const lock = LockService.getScriptLock();


  // รอสิทธิ์บันทึกไม่เกิน 30 วินาที

  if (!lock.tryLock(30000))

    throw Error('มีการบันทึกพร้อมกัน กรุณาลองอีกครั้ง');


  const undo = [];


  try {

    const db = db_(),

      t = table_(db, kind),

      pk = t.headers[0],

      id = ({ customer: 'CUS_', customer_contact: 'CON_', reports: 'RPT_' })[kind] + token;


    // หากรหัสนี้มีแล้ว คืนผลทันทีโดยไม่เพิ่มแถว

    if (t.rows.some(r => String(r[pk]) === id)) return { id, duplicate: true };


    // obj เก็บรายการหลัก; team เก็บรหัสทีมงาน; join เก็บตารางเชื่อมรายงานกับทีมงาน

    let obj = { [pk]: id }, team = [], join;


    if (kind === 'customer') {

      // ----- ลูกค้า: ตรวจชื่อซ้ำ ประเภท สถานะ และข้อมูลประกอบ -----

      obj.company_name = text_(p, 'company_name', true, 250);

      obj.company_short_name = text_(p, 'company_short_name', false, 80);


      if (t.rows.some(r =>

        String(r.company_name).trim().toLowerCase() === obj.company_name.toLowerCase()

      )) throw Error('มีชื่อลูกค้านี้แล้ว กรุณาตรวจสอบรายการเดิม');


      obj.customer_type = option_(p, 'customer_type', ['CUSTOMER', 'PROSPECT']);

      obj.status = option_(p, 'status', ['ACTIVE', 'INACTIVE']);

      ['phone', 'address', 'province', 'remark'].forEach(k => obj[k] = text_(p, k, false));

      obj.location_url = url_(p, 'location_url');

    } else {

      // ทั้งผู้ติดต่อและรายงานต้องอ้างอิงลูกค้าที่ใช้งานได้

      obj.customer_id = text_(p, 'customer_id', true);

      ref_(table_(db, 'customer').rows, 'customer_id', obj.customer_id, 'ลูกค้า', true);


      if (kind === 'customer_contact') {

        // ----- ผู้ติดต่อ: บังคับชื่อ ตรวจอีเมล และชื่อ-นามสกุลซ้ำในบริษัทเดียวกัน -----

        ['ชื่อ', 'นามสกุล', 'ตำแหน่ง', 'โทรศัพท์', 'email'].forEach(k =>

          obj[k] = text_(p, k, k === 'ชื่อ', k === 'email' ? 254 : 250)

        );


        if (obj.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(obj.email))

          throw Error('รูปแบบอีเมลไม่ถูกต้อง');


        if (t.rows.some(r =>

          r.customer_id === obj.customer_id &&

          String(r['ชื่อ']).trim() === obj['ชื่อ'] &&

          String(r['นามสกุล']).trim() === obj['นามสกุล']

        )) throw Error('มีผู้ติดต่อชื่อนี้ในบริษัทแล้ว');

      } else {

        // ----- รายงาน: ตรวจผู้ติดต่อ วันที่ กิจกรรม ผู้บันทึก และทีมงาน -----

        obj.contact_id = text_(p, 'contact_id', true);

        const c = ref_(

          table_(db, 'customer_contact').rows,

          'contact_id', obj.contact_id, 'ผู้ติดต่อ'

        );


        if (c.customer_id !== obj.customer_id)

          throw Error('ผู้ติดต่อไม่ได้อยู่ในบริษัทที่เลือก');


        obj.visit_at = serial_(text_(p, 'visit_at', true), true);

        obj.next_follow_up = p.next_follow_up ? serial_(p.next_follow_up, false) : '';


        // ตัดเวลาออกด้วย Math.floor เพื่อให้วันติดตามเป็นวันเดียวกับวันปฏิบัติงานได้

        if (obj.next_follow_up !== '' && obj.next_follow_up < Math.floor(obj.visit_at))

          throw Error('วันติดตามต้องไม่ก่อนวันปฏิบัติงาน');


        obj.activity = text_(p, 'activity', true, 150);

        obj.job_detail = text_(p, 'job_detail', true, 10000);

        obj.remark = text_(p, 'remark', false, 5000);

        obj.status = option_(p, 'status', ['PLANNED', 'IN_PROGRESS', 'COMPLETED', 'CANCELLED']);

        obj.attachment_url = url_(p, 'attachment_url');


        const employees = table_(db, 'employee').rows;

        obj.created_by = text_(p, 'created_by', true);

        ref_(employees, 'employee_id', obj.created_by, 'ผู้บันทึก', true);


        if (!Array.isArray(p.employee_ids) || !p.employee_ids.length)

          throw Error('กรุณาเลือกผู้ปฏิบัติงานอย่างน้อย 1 คน');


        // แปลงรหัสเป็นข้อความ ตัดรหัสซ้ำ และตรวจว่าพนักงานแต่ละคนใช้งานได้

        team = [...new Set(p.employee_ids.map(String))];

        team.forEach(e => ref_(employees, 'employee_id', e, 'พนักงาน', true));

        join = table_(db, 'report_employee');


        // เวลาเริ่มต้นทั้งสองช่องเท่ากัน โดยใช้เวลาประเทศไทย ความละเอียดระดับนาที

        obj.created_at = serial_(Utilities.formatDate(new Date(), GD.tz, "yyyy-MM-dd'T'HH:mm"), true);

        obj.updated_at = obj.created_at;

      }

    }


    // เขียนรายการหลักเมื่อผ่านการตรวจสอบข้างต้นแล้ว

    const row = append_(t, obj, undo);


    if (kind === 'reports') {

      // รายงานหนึ่งรายการเชื่อมพนักงานหลายคนได้ แต่ละคนมีแถวเชื่อมของตนเอง

      team.forEach((employee_id, i) => append_(join, {

        report_id: id,

        employee_id,

        report_employee_id: 'RPE_' + token + '_' + (i + 1)

      }, undo));


      // คอลัมน์ 2=visit_at, 12=created_at, 13=updated_at และ 9=next_follow_up

      [2, 12, 13].forEach(col => t.sheet.getRange(row, col).setNumberFormat('yyyy-mm-dd hh:mm'));

      t.sheet.getRange(row, 9).setNumberFormat('yyyy-mm-dd');

    }


    SpreadsheetApp.flush();

    return { id, duplicate: false };

  } catch (e) {

    // ย้อนล้างเนื้อหาช่วงเซลล์ที่ลงทะเบียนไว้ โดยเริ่มจากช่วงล่าสุด

    // clearContent ไม่ลบรูปแบบเซลล์หรือแถวที่ insertRowsAfter เพิ่มไว้

    if (undo.length) {

      try {

        undo.reverse().forEach(r => r.clearContent());

        SpreadsheetApp.flush();

      } catch (rollback) {

        throw Error('บันทึกไม่สมบูรณ์และคืนค่าล้มเหลว กรุณาตรวจสอบตารางก่อนส่งอีกครั้ง: ' + e.message);

      }

    }


    throw e;

  } finally {

    // ปลด lock เสมอเมื่อออกจาก try ไม่ว่าจะสำเร็จ คืนผลซ้ำ หรือเกิดข้อผิดพลาด

    lock.releaseLock();

  }

}


// ============================================================================

// 6. ตรวจสอบความสัมพันธ์ของข้อมูล (อ่านอย่างเดียว)

// ============================================================================


/** Read-only integrity report. No test data is inserted. */

/**

 * ตรวจรหัสซ้ำและความสัมพันธ์ที่กำหนดไว้ในโค้ดเดิม

 * แสดงจำนวนรายการแต่ละตารางและ issues ใน execution log แล้วคืนอาร์เรย์ issues

 * [] หมายถึงไม่พบปัญหาตามเงื่อนไขที่ฟังก์ชันนี้ตรวจ ไม่ได้ยืนยันข้อมูลทุกด้าน

 * ฟังก์ชันนี้ไม่แก้ไขข้อมูล และไม่ได้ตรวจรหัส report_employee_id ซ้ำโดยตรง

 */

function auditData() {

  const d = getBootstrap(), issues = [];


  // ตรวจ primary key ซ้ำในสี่ตารางหลัก

  ['customer', 'customer_contact', 'reports', 'employee'].forEach(n => {

    const key = HEADERS[n][0], seen = new Set();

    d[n].forEach(r => {

      if (seen.has(r[key])) issues.push(n + ': รหัสซ้ำ ' + r[key]);

      seen.add(r[key]);

    });

  });


  // เตรียมชุดรหัสสำหรับตรวจการอ้างอิงข้ามตาราง

  const customers = new Set(d.customer.map(r => r.customer_id)),

    employees = new Set(d.employee.map(r => r.employee_id)),

    reports = new Set(d.reports.map(r => r.report_id));


  d.customer_contact.forEach(c => {

    if (!customers.has(c.customer_id))

      issues.push('ผู้ติดต่อไม่มีลูกค้า: ' + c.contact_id);

  });


  d.reports.forEach(r => {

    const c = d.customer_contact.find(c => c.contact_id === r.contact_id);

    if (!customers.has(r.customer_id) || !c || c.customer_id !== r.customer_id)

      issues.push('ลูกค้า/ผู้ติดต่อไม่สัมพันธ์กัน: ' + r.report_id);

    if (!employees.has(r.created_by))

      issues.push('ผู้บันทึกไม่อยู่ใน employee: ' + r.report_id);

  });


  // ตรวจคู่รายงาน/พนักงานซ้ำ และการอ้างอิงรายงานหรือพนักงานที่ไม่มีอยู่

  const pairs = new Set();

  d.report_employee.forEach(r => {

    const pair = r.report_id + '|' + r.employee_id;

    if (pairs.has(pair)) issues.push('ทีมงานซ้ำ: ' + pair);

    pairs.add(pair);

    if (!reports.has(r.report_id) || !employees.has(r.employee_id))

      issues.push('ทีมงานอ้างอิงข้อมูลที่ไม่มี: ' + pair);

  });


  console.log(JSON.stringify({

    counts: Object.fromEntries(Object.keys(HEADERS).map(n => [n, d[n].length])),

    issues

  }, null, 2));


  return issues;

}


// ============================================================================

// 7. แก้ไขรายการเดิม: ตรวจข้อมูลก่อนเขียน และคืนค่าเดิมเมื่อเขียนไม่สำเร็จ

// ============================================================================

/** expected เป็นข้อมูลตอนเปิดฟอร์ม ใช้ตรวจว่ามีคนอื่นแก้รายการนี้ไปแล้วหรือไม่ */

function updateRecord(kind, id, p, expected, expectedTeam) {

  if (!['customer','customer_contact','reports'].includes(kind)) throw Error('ประเภทข้อมูลไม่ถูกต้อง');

  id=String(id || ''); p=p||{};

  if(!id || !expected) throw Error('ไม่พบข้อมูลต้นฉบับ กรุณาโหลดข้อมูลใหม่');

  const lock=LockService.getScriptLock();

  if(!lock.tryLock(30000)) throw Error('มีการบันทึกพร้อมกัน กรุณาลองอีกครั้ง');

  const backups=[];

  try {

    const db=db_(), t=table_(db,kind), pk=t.headers[0];

    const matches=t.rows.filter(r=>String(r[pk])===id);

    if(matches.length!==1) throw Error('ไม่พบรายการหรือมีรหัสซ้ำ กรุณาตรวจสอบข้อมูล');

    const original=matches[0], tz=db.getSpreadsheetTimeZone();

    if(t.headers.filter(h=>h!=='visit_at_th').some(h=>display_(original[h],tz)!==String(expected[h] == null?'':expected[h])))

      throw Error('ข้อมูลถูกแก้ไขหลังจากเปิดฟอร์ม กรุณาโหลดข้อมูลใหม่ก่อนแก้ไข');

    const oldJoin=kind==='reports'?table_(db,'report_employee'):null;

    const oldTeam=oldJoin?[...new Set(oldJoin.rows.filter(r=>String(r.report_id)===id).map(r=>String(r.employee_id)))]:[];

    if(kind==='reports' && JSON.stringify(oldTeam.slice().sort())!==JSON.stringify([...new Set((expectedTeam||[]).map(String))].sort()))

      throw Error('ทีมงานถูกแก้ไขแล้ว กรุณาโหลดข้อมูลใหม่');

    // ตรวจชื่อซ้ำโดยไม่นับรายการที่กำลังแก้ไขเอง

    t.rows=t.rows.filter(r=>String(r[pk])!==id);

    let obj={ [pk]:id }, team=[], join;

    if (kind === 'customer') {

      // ----- ลูกค้า: ตรวจชื่อซ้ำ ประเภท สถานะ และข้อมูลประกอบ -----

      obj.company_name = text_(p, 'company_name', true, 250);

      obj.company_short_name = text_(p, 'company_short_name', false, 80);


      if (t.rows.some(r =>

        String(r.company_name).trim().toLowerCase() === obj.company_name.toLowerCase()

      )) throw Error('มีชื่อลูกค้านี้แล้ว กรุณาตรวจสอบรายการเดิม');


      obj.customer_type = option_(p, 'customer_type', ['CUSTOMER', 'PROSPECT']);

      obj.status = option_(p, 'status', ['ACTIVE', 'INACTIVE']);

      ['phone', 'address', 'province', 'remark'].forEach(k => obj[k] = text_(p, k, false));

      obj.location_url = url_(p, 'location_url');

    } else {

      // ทั้งผู้ติดต่อและรายงานต้องอ้างอิงลูกค้าที่ใช้งานได้

      obj.customer_id = text_(p, 'customer_id', true);

      ref_(table_(db, 'customer').rows, 'customer_id', obj.customer_id, 'ลูกค้า', obj.customer_id !== String(original.customer_id));


      if (kind === 'customer_contact') {

        // ----- ผู้ติดต่อ: บังคับชื่อ ตรวจอีเมล และชื่อ-นามสกุลซ้ำในบริษัทเดียวกัน -----

        ['ชื่อ', 'นามสกุล', 'ตำแหน่ง', 'โทรศัพท์', 'email'].forEach(k =>

          obj[k] = text_(p, k, k === 'ชื่อ', k === 'email' ? 254 : 250)

        );


        if (obj.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(obj.email))

          throw Error('รูปแบบอีเมลไม่ถูกต้อง');


        if (t.rows.some(r =>

          r.customer_id === obj.customer_id &&

          String(r['ชื่อ']).trim() === obj['ชื่อ'] &&

          String(r['นามสกุล']).trim() === obj['นามสกุล']

        )) throw Error('มีผู้ติดต่อชื่อนี้ในบริษัทแล้ว');

      } else {

        // ----- รายงาน: ตรวจผู้ติดต่อ วันที่ กิจกรรม ผู้บันทึก และทีมงาน -----

        obj.contact_id = text_(p, 'contact_id', true);

        const c = ref_(

          table_(db, 'customer_contact').rows,

          'contact_id', obj.contact_id, 'ผู้ติดต่อ'

        );


        if (c.customer_id !== obj.customer_id)

          throw Error('ผู้ติดต่อไม่ได้อยู่ในบริษัทที่เลือก');


        obj.visit_at = serial_(text_(p, 'visit_at', true), true);

        obj.next_follow_up = p.next_follow_up ? serial_(p.next_follow_up, false) : '';


        // ตัดเวลาออกด้วย Math.floor เพื่อให้วันติดตามเป็นวันเดียวกับวันปฏิบัติงานได้

        if (obj.next_follow_up !== '' && obj.next_follow_up < Math.floor(obj.visit_at))

          throw Error('วันติดตามต้องไม่ก่อนวันปฏิบัติงาน');


        obj.activity = text_(p, 'activity', true, 150);

        obj.job_detail = text_(p, 'job_detail', true, 10000);

        obj.remark = text_(p, 'remark', false, 5000);

        obj.status = option_(p, 'status', ['PLANNED', 'IN_PROGRESS', 'COMPLETED', 'CANCELLED']);

        obj.attachment_url = url_(p, 'attachment_url');


        const employees = table_(db, 'employee').rows;

        obj.created_by = text_(p, 'created_by', true);

        ref_(employees, 'employee_id', obj.created_by, 'ผู้บันทึก', obj.created_by !== String(original.created_by));


        if (!Array.isArray(p.employee_ids) || !p.employee_ids.length)

          throw Error('กรุณาเลือกผู้ปฏิบัติงานอย่างน้อย 1 คน');


        // แปลงรหัสเป็นข้อความ ตัดรหัสซ้ำ และตรวจว่าพนักงานแต่ละคนใช้งานได้

        team = [...new Set(p.employee_ids.map(String))];

        team.forEach(e => ref_(employees, 'employee_id', e, 'พนักงาน', !oldTeam.includes(e)));

        join = table_(db, 'report_employee');


        // เวลาเริ่มต้นทั้งสองช่องเท่ากัน โดยใช้เวลาประเทศไทย ความละเอียดระดับนาที

        obj.created_at = serial_(Utilities.formatDate(new Date(), GD.tz, "yyyy-MM-dd'T'HH:mm"), true);

        obj.updated_at = obj.created_at;

      }

    }


    // ห้ามย้ายผู้ติดต่อไปบริษัทอื่นจนรายงานเดิมอ้างอิงผิดบริษัท

    if(kind==='customer_contact' && String(original.customer_id)!==obj.customer_id && table_(db,'reports').rows.some(r=>String(r.contact_id)===id))

      throw Error('ผู้ติดต่อนี้มีรายงานอ้างอิงอยู่ จึงย้ายบริษัทไม่ได้');

    if(kind==='reports') obj.created_at=original.created_at;

    const ids=t.sheet.getRange(1,1,t.last,1).getValues();

    const row=ids.findIndex((r,i)=>i>0 && String(r[0])===id)+1;

    if(row<2) throw Error('ไม่พบแถวต้นฉบับ');

    const write=(range, values)=>{

      const previous=range.getValues(), formulas=range.getFormulas();

      backups.push({range,values:previous.map((r,i)=>r.map((v,j)=>formulas[i][j]||v))});

      range.setValues(values);

    };

    const width=t.headers.filter(h=>h!=='visit_at_th').length;

    // ไม่แตะคอลัมน์สูตร visit_at_th หรือคอลัมน์เพิ่มเติมนอกโครงสร้างเดิม

    write(t.sheet.getRange(row,1,1,width),[t.headers.slice(0,width).map(h=>safe_(obj[h]===undefined?'':obj[h]))]);

    if(kind==='reports') {

      const rows=join.last>1?join.sheet.getRange(2,1,join.last-1,3).getValues():[];

      const retained=new Set();

      rows.forEach((r,i)=>{

        if(String(r[0])!==id) return;

        const employee=String(r[1]);

        if(team.includes(employee)&&!retained.has(employee)){retained.add(employee);return;}

        // ล้างเฉพาะความสัมพันธ์ที่ยกเลิกหรือซ้ำ ไม่ลบแถวและไม่เลื่อนข้อมูลอื่น

        write(join.sheet.getRange(i+2,1,1,3),[['','','']]);

      });

      team.filter(e=>!retained.has(e)).forEach(employee=>{

        const next=++join.last;

        if(next>join.sheet.getMaxRows())join.sheet.insertRowsAfter(join.sheet.getMaxRows(),100);

        write(join.sheet.getRange(next,1,1,3),[[id,employee,'RPE_'+Utilities.getUuid().replace(/-/g,'')]]);

      });

    }

    SpreadsheetApp.flush();

    return {id,updated:true};

  } catch(error) {

    try {backups.reverse().forEach(b=>b.range.setValues(b.values));SpreadsheetApp.flush();}

    catch(rollback){throw Error('คืนค่าหลังแก้ไขไม่สำเร็จ กรุณาตรวจสอบตาราง: '+error.message);}

    throw error;

  } finally {lock.releaseLock();}

}
