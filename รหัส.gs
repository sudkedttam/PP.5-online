function doGet(e) {  
  return HtmlService.createTemplateFromFile('home').evaluate()
        .setTitle("อ32102-811-2565-2")
        .addMetaTag('viewport','width=device-width , initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  } 
  
  function getCode(code) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allss =ss.getSheets();  
  for (var i in allss){
    var ws =ss.getSheets()[i];
    var data=ws.getDataRange().getDisplayValues().filter(row=>{
      return row[1]==code
      })
      Logger.log(data)
    if(data.length>1) break;
  }
  
  var stdCodesList = data.map (function(r) { return r[1]; }); 
  
  var stdList = data.map(function(r) { 
    
  return [`<table class="table table-bordered mt-3">  
          <thead class="p-3 mb-2 bd-yellow-500 text-black">
          <tr>
          <th scope="col" colspan="12"><center><i class="fas fa-user-graduate"></i> ข้อมูลผู้เข้าระบบ</center></th>
          </tr>
          </thead>
          <tbody>
          <!--Start Personal Data-->
          <tr>
          <td colspan="12"><center><b>รหัสเข้าระบบ :</b> ${r[1]}</center></td>
          </tr> 
          <tr>
          <td colspan="12"><center><b>ชื่อ-สกุล :</b> ${r[2]}</center></td>
          </tr> 
          <tr>
          <td colspan="12"><center><b>ห้อง/เลขที่ :</b> ${r[3]}</center></td>
          </tr>
          <tr>
          <td colspan="12"><center><b>สถานะ :</b> ${r[152]}</center></td>
          </tr> 
          <tr>
          <td colspan="12"><center><b>ร้อยละเวลาเรียน :</b> ${r[153]}</center></td>
          </tr> 
          <tr>
          <td colspan="12"><center><b>รวมไม่เข้าเรียน (ครั้ง) :</b> ${r[154]}</center></td>
          </tr>
          <tr>
          <td colspan="12"><center><b>ขาด (ครั้ง) :</b> ${r[155]}</center></td>
          </tr> 
          <tr>
          <td colspan="12"><center><b>ลาป่วย (ครั้ง) :</b> ${r[156]}</center></td>
          </tr> 
          <tr>
          <td colspan="12"><center><b>ลากิจ (ครั้ง) :</b> ${r[157]}</center></td>
          </tr>
          <tr>
          <td colspan="12"><center><b>หลบ (ครั้ง) :</b> ${r[158]}</center></td>
          </tr> 
          <!--End Personal Data-->

          <thead class="p-3 mb-2 bd-yellow-500 text-black">
          <tr>
          <th scope="col" ><center>ที่มาคะแนน</center></th>
          <th scope="col"><center>ได้/เต็ม</center></th>
          </tr>
          </thead>

    <!...start...S1>
    <tr>
      <td>
        <left><b>${r[4]}</b></left>
      </td>
      <td><center>${r[5]}</center></td>
    </tr>
    <!...end...S1>

    <!...start...S2>
    <tr>
      <td>
        <left><b>${r[6]}</b></left>
      </td>
      <td><center>${r[7]}</center></td>
    </tr>
    <!...end...S2>

    <!...start...S3>
    <tr>
      <td>
        <left><b>${r[8]}</b></left>
      </td>
      <td><center>${r[9]}</center></td>
    </tr>
    <!...end...S3>

    <!...start...S4>
    <tr>
      <td>
        <left><b>${r[10]}</b></left>
      </td>
      <td><center>${r[11]}</center></td>
    </tr>
    <!...end...S4>

    <!...start...S5>
    <tr>
      <td>
        <left><b>${r[12]}</b></left>
      </td>
      <td><center>${r[13]}</center></td>
    </tr>
    <!...end...S5>

    <!...start...S6>
    <tr>
      <td>
        <left><b>${r[14]}</b></left>
      </td>
      <td><center>${r[15]}</center></td>
    </tr>
    <!...end...S6>

    <!...start...S7>
    <tr>
      <td>
        <left><b>${r[16]}</b></left>
      </td>
      <td><center>${r[17]}</center></td>
    </tr>
    <!...end...S7>

    <!...start...S8>
    <tr>
      <td>
        <left><b>${r[18]}</b></left>
      </td>
      <td><center>${r[19]}</center></td>
    </tr>
    <!...end...S8>

    <!...start...S9>
    <tr>
      <td>
        <left><b>${r[20]}</b></left>
      </td>
      <td><center>${r[21]}</center></td>
    </tr>
    <!...end...S9>

    <!...start...S10>
    <tr>
      <td>
        <left><b>${r[22]}</b></left>
      </td>
      <td><center>${r[23]}</center></td>
    </tr>
    <!...end...S10>

    <!...start...S11>
    <tr>
      <td>
        <left><b>${r[24]}</b></left>
      </td>
      <td><center>${r[25]}</center></td>
    </tr>
    <!...end...S11>

    <!...start...S12>
    <tr>
      <td>
        <left><b>${r[26]}</b></left>
      </td>
      <td><center>${r[27]}</center></td>
    </tr>
    <!...end...S12>

    <!...start...S13>
    <tr>
      <td>
        <cenlefter><b>${r[28]}</b></left>
      </td>
      <td><center>${r[29]}</center></td>
    </tr>
    <!...end...S13>

    <!...start...S14>
    <tr>
      <td>
        <left><b>${r[30]}</b></left>
      </td>
      <td><center>${r[31]}</center></td>
    </tr>
    <!...end...S14>

    <!...start...S15>
    <tr>
      <td>
        <left><b>${r[32]}</b></left>
      </td>
      <td><center>${r[33]}</center></td>
    </tr>
    <!...end...S15>

    <!...start...S16>
    <tr>
      <td>
        <left><b>${r[34]}</b></left>
      </td>
      <td><center>${r[35]}</center></td>
    </tr>
    <!...end...S16>

    <!...start...S17>
    <tr>
      <td>
        <left><b>${r[36]}</b></left>
      </td>
      <td><center>${r[37]}</center></td>
    </tr>
    <!...end...S17>

    <!...start...S18>
    <tr>
      <td>
        <left><b>${r[38]}</b></left>
      </td>
      <td><center>${r[39]}</center></td>
    </tr>
    <!...end...S18>

    <!...start...S19>
    <tr>
      <td>
        <left><b>${r[40]}</b></left>
      </td>
      <td><center>${r[41]}</center></td>
    </tr>
    <!...end...S19>

    <!...start...S20>
    <tr>
      <td>
        <left><b>${r[42]}</b></left>
      </td>
      <td><center>${r[43]}</center></td>
    </tr>
    <!...end...S20>

    <!...start...S21>
    <tr>
      <td>
        <left><b>${r[44]}</b></left>
      </td>
      <td><center>${r[45]}</center></td>
    </tr>
    <!...end...S21>

    <!...start...S22>
    <tr>
      <td>
        <left><b>${r[46]}</b></left>
      </td>
      <td><center>${r[47]}</center></td>
    </tr>
    <!...end...S22>

    <!...start...S23>
    <tr>
      <td>
        <left><b>${r[48]}</b></left>
      </td>
      <td><center>${r[49]}</center></td>
    </tr>
    <!...end...S23>

    <!...start...S24>
    <tr>
      <td>
        <left><b>${r[50]}</b></left>
      </td>
      <td><center>${r[51]}</center></td>
    </tr>
    <!...end...S24>

    <!...start...S25>
    <tr>
      <td>
        <left><b>${r[52]}</b></left>
      </td>
      <td><center>${r[53]}</center></td>
    </tr>
    <!...end...S25>

    <!...start...S26>
    <tr>
      <td>
        <left><b>${r[54]}</b></left>
      </td>
      <td><center>${r[55]}</center></td>
    </tr>
    <!...end...S26>

    <!...start...S27>
    <tr>
      <td>
        <left><b>${r[56]}</b></left>
      </td>
      <td><center>${r[57]}</center></td>
    </tr>
    <!...end...S27>

    <!...start...S28>
    <tr>
      <td>
        <left><b>${r[58]}</b></left>
      </td>
      <td><center>${r[59]}</center></td>
    </tr>
    <!...end...S28>

    <!...start...S29>
    <tr>
      <td>
        <left><b>${r[60]}</b></left>
      </td>
      <td><center>${r[61]}</center></td>
    </tr>
    <!...end...S29>

    <!...start...S30>
    <tr>
      <td>
        <left><b>${r[62]}</b></left>
      </td>
      <td><center>${r[63]}</center></td>
    </tr>
    <!...end...S30>

    <!...start...S31>
    <tr>
      <td>
        <left><b>${r[64]}</b></left>
      </td>
      <td><center>${r[65]}</center></td>
    </tr>
    <!...end...S31>

    <!...start...S32>
    <tr>
      <td>
        <left><b>${r[66]}</b></left>
      </td>
      <td><center>${r[67]}</center></td>
    </tr>
    <!...end...S32>

    <!...start...S33>
    <tr>
      <td>
        <left><b>${r[68]}</b></left>
      </td>
      <td><center>${r[69]}</center></td>
    </tr>
    <!...end...S33>

    <!...start...S34>
    <tr>
      <td>
        <left><b>${r[70]}</b></left>
      </td>
      <td><center>${r[71]}</center></td>
    </tr>
    <!...end...S34>

    <!...start...S35>
    <tr>
      <td>
        <left><b>${r[72]}</b></left>
      </td>
      <td><center>${r[73]}</center></td>
    </tr>
    <!...end...S35>

    <!...start...S36>
    <tr>
      <td>
        <left><b>${r[74]}</b></left>
      </td>
      <td><center>${r[75]}</center></td>
    </tr>
    <!...end...S36>

    <!...start...S37>
    <tr>
      <td>
        <left><b>${r[76]}</b></left>
      </td>
      <td><center>${r[77]}</center></td>
    </tr>
    <!...end...S37>

    <!...start...S38>
    <tr>
      <td>
        <left><b>${r[78]}</b></left>
      </td>
      <td><center>${r[79]}</center></td>
    </tr>
    <!...end...S38>

    <!...start...S39>
    <tr>
      <td>
        <left><b>${r[80]}</b></left>
      </td>
      <td><center>${r[81]}</center></td>
    </tr>
    <!...end...S39>

    <!...start...S40>
    <tr>
      <td>
        <left><b>${r[82]}</b></left>
      </td>
      <td><center>${r[83]}</center></td>
    </tr>
    <!...end...S40>

    <!...start...S41>
    <tr>
      <td>
        <left><b>${r[84]}</b></left>
      </td>
      <td><center>${r[85]}</center></td>
    </tr>
    <!...end...S41>

    <!...start...S42>
    <tr>
      <td>
        <left><b>${r[86]}</b></left>
      </td>
      <td><center>${r[87]}</center></td>
    </tr>
    <!...end...S42>

    <!...start...S43>
    <tr>
      <td>
        <left><b>${r[88]}</b></left>
      </td>
      <td><center>${r[89]}</center></td>
    </tr>
    <!...end...S43>

    <!...start...S44>
    <tr>
      <td>
        <left><b>${r[90]}</b></left>
      </td>
      <td><center>${r[91]}</center></td>
    </tr>
    <!...end...S44>

    <!...start...S45>
    <tr>
      <td>
        <left><b>${r[92]}</b></left>
      </td>
      <td><center>${r[93]}</center></td>
    </tr>
    <!...end...S45>

    <!...start...S46>
    <tr>
      <td>
        <left><b>${r[94]}</b></left>
      </td>
      <td><center>${r[95]}</center></td>
    </tr>
    <!...end...S46>

    <!...start...S47>
    <tr>
      <td>
        <left><b>${r[96]}</b></left>
      </td>
      <td><center>${r[97]}</center></td>
    </tr>
    <!...end...S47>

    <!...start...S48>
    <tr>
      <td>
        <left><b>${r[98]}</b></left>
      </td>
      <td><center>${r[99]}</center></td>
    </tr>
    <!...end...S48>

    <!...start...S49>
    <tr>
      <td>
        <left><b>${r[100]}</b></left>
      </td>
      <td><center>${r[101]}</center></td>
    </tr>
    <!...end...S49>

    <!...start...S50>
    <tr>
      <td>
        <left><b>${r[102]}</b></left>
      </td>
      <td><center>${r[103]}</center></td>
    </tr>
    <!...end...S50>

    <!...start...S51>
    <tr>
      <td>
        <left><b>${r[104]}</b></left>
      </td>
      <td><center>${r[105]}</center></td>
    </tr>
    <!...end...S51>

    <!...start...S52>
    <tr>
      <td>
        <left><b>${r[106]}</b></left>
      </td>
      <td><center>${r[107]}</center></td>
    </tr>
    <!...end...S52>

    <!...start...S53>
    <tr>
      <td>
        <left><b>${r[108]}</b></left>
      </td>
      <td><center>${r[109]}</center></td>
    </tr>
    <!...end...S53>

    <!...start...S54>
    <tr>
      <td>
        <left><b>${r[110]}</b></left>
      </td>
      <td><center>${r[111]}</center></td>
    </tr>
    <!...end...S54>

    <!...start...S55>
    <tr>
      <td>
        <left><b>${r[112]}</b></left>
      </td>
      <td><center>${r[113]}</center></td>
    </tr>
    <!...end...S55>

    <!...start...S56>
    <tr>
      <td>
        <left><b>${r[114]}</b></left>
      </td>
      <td><center>${r[115]}</center></td>
    </tr>
    <!...end...S56>

    <!...start...S57>
    <tr>
      <td>
        <left><b>${r[116]}</b></left>
      </td>
      <td><center>${r[117]}</center></td>
    </tr>
    <!...end...S57>

    <!...start...S58>
    <tr>
      <td>
        <left><b>${r[118]}</b></left>
      </td>
      <td><center>${r[119]}</center></td>
    </tr>
    <!...end...S58>

    <!...start...S59>
    <tr>
      <td>
        <left><b>${r[120]}</b></left>
      </td>
      <td><center>${r[121]}</center></td>
    </tr>
    <!...end...S59>

    <!...start...S60>
    <tr>
      <td>
        <left><b>${r[122]}</b></left>
      </td>
      <td><center>${r[123]}</center></td>
    </tr>
    <!...end...S60>

    <!...start...S61>
    <tr>
      <td>
        <left><b>${r[124]}</b></left>
      </td>
      <td><center>${r[125]}</center></td>
    </tr>
    <!...end...S61>

    <!...start...S62>
    <tr>
      <td>
        <left><b>${r[126]}</b></left>
      </td>
      <td><center>${r[127]}</center></td>
    </tr>
    <!...end...S62>

    <!...start...S63>
    <tr>
      <td>
        <left><b>${r[128]}</b></left>
      </td>
      <td><center>${r[129]}</center></td>
    </tr>
    <!...end...S63>

    <!...start...S64>
    <tr>
      <td>
        <left><b>${r[130]}</b></left>
      </td>
      <td><center>${r[131]}</center></td>
    </tr>
    <!...start...S64>

    <!...start...S65>
    <tr>
      <td>
        <left><b>${r[132]}</b></left>
      </td>
      <td><center>${r[133]}</center></td>
    </tr>
    <!...end...S65>

    <!...start...S66>
    <tr>
      <td>
        <left><b>${r[134]}</b></left>
      </td>
      <td><center>${r[135]}</center></td>
    </tr>
    <!...end...S66>

    <!...start...S67>
    <tr>
      <td>
        <left><b>${r[136]}</b></left>
      </td>
      <td><center>${r[137]}</center></td>
    </tr>
    <!...end...S67>

    <!...start...S68>
    <tr>
      <td>
        <left><b>${r[138]}</b></left>
      </td>
      <td><center>${r[139]}</center></td>
    </tr>
    <!...end...S68>

    <!...start...S69>
    <tr>
      <td>
        <left><b>${r[140]}</b></left>
      </td>
      <td><center>${r[141]}</center></td>
    </tr>
    <!...end...S60>

    <!...start...S70>
    <tr>
      <td>
        <left><b>${r[142]}</b></left>
      </td>
      <td><center>${r[143]}</center></td>
    </tr>
    <!...end...S70>

    <!...start...S71>
    <tr>
      <td>
        <left><b>${r[144]}</b></left>
      </td>
      <td><center>${r[145]}</center></td>
    </tr>
    <!...end...S71>

    <!...start...S72>
    <tr>
      <td>
        <left><b>${r[146]}</b></left>
      </td>
      <td><center>${r[147]}</center></td>
    </tr>
    <!...end...S72>

    <!...start...S73>
    <tr>
      <td>
        <left><b>${r[148]}</b></left>
      </td>
      <td><center>${r[149]}</center></td>
    </tr>
    <!...end...S73>

    <!...start...S74>
    <tr>
      <td>
        <left><b>${r[150]}</b></left>
      </td>
      <td><center>${r[151]}</center></td>
    </tr>
    <!...end...S74>

          </tbody>
          </table>                   
          `];
  });
  
  var position = stdCodesList.indexOf(code); 
  if(position > -1){
  return stdList[position];
  } else {
  return '<center>*ไม่พบรหัสนี้ขอรับ*<br><img src="https://drive.google.com/uc?id=1tu2D2viPPVc5Bzu_SQHKzPvS1W60ZJrc" width="200" height="200"></center>';
    } 
  }
  
  function getURL(){
  return ScriptApp.getService().getUrl()
  }
