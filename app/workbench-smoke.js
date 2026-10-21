// Real Electron + Python integration checks, isolated from the user's settings.
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

exports.run = async (win, root) => {
  fs.mkdirSync(path.join(root, 'tmp'), {recursive: true});
  const out = fs.mkdtempSync(path.join(root, 'tmp', 'workbench-ui-'));
  const base = path.join(root, '测试集', '泛用性验收_v3', 'W1');
  const source = path.join(base, 'W1_原始计划表.xlsx');
  const shifted = path.join(out, '表头移位与列换序.xlsx');
  const duplicate = path.join(out, '重复字段.xlsx');
  const compatibility = path.join(out, '扩展样式输入.xlsx');
  const sameBook = path.join(out, '同簿重复字段.xlsx');
  const otherMapped = path.join(out, '另一来源范围.xlsx');
  const normalizedA = path.join(out, '规范字段甲.xlsx');
  const normalizedB = path.join(out, '规范字段乙.xlsx');
  const uncertain = path.join(out, '识别方向待确认.xlsx');
  const driftTemplate = path.join(out, '外部换列模板.xlsx');
  const driftSource = path.join(out, '外部换列来源.xlsx');
  const errors = [];
  win.webContents.on('console-message', (_event, level, message) => {
    if (level >= 3 && !/GPU|EGL|ANGLE|GLES/i.test(message)) errors.push(message);
  });
  execFileSync(process.env.TOOLBOX_PYTHON || 'python', ['-c', `
import sys
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl import Workbook, load_workbook
src=load_workbook(sys.argv[1],data_only=True).active
book=Workbook(); ws=book.active; ws.title='移位结构'
ws['A1']='通用结构复用验收'; ws.merge_cells('A1:G1')
order=[7,4,1,5,2,3,6]
for row in range(4,17):
    for col, old in enumerate(order,1): ws.cell(row+3,col,src.cell(row,old).value)
for col,old in enumerate(order,1): ws.cell(20,col,src.cell(5,old).value)
ws.cell(20,3,'TNEW')
book.save(sys.argv[2])
ws.cell(7,5,ws.cell(7,6).value)
book.save(sys.argv[3])
compat=Workbook(); sheet=compat.active; sheet.title='样式兼容验收'
sheet.append(['编号','分类','数值','计算']); sheet.append(['X001','甲',12,'=C2*2']); sheet.append(['X002','乙',34,'=SUM(C2:C3)'])
buffer=BytesIO(); compat.save(buffer)
extension=b'<extLst><ext uri="{C66500BD-6782-4325-8C5A-71D048A78866}"><demo:format xmlns:demo="urn:workbench-smoke" value="1"/></ext></extLst>'
workbook_extension=b'<extLst><ext uri="{79F54976-1DA5-4618-B147-4CDE4B953A38}"><x15:absPath xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" url="file:///C:/synthetic-workbench-fixture/"/></ext><ext uri="urn:workbench-smoke:unknown"><demo:futureFeature xmlns:demo="urn:workbench-smoke" value="1"/></ext></extLst>'
with ZipFile(BytesIO(buffer.getvalue())) as original, ZipFile(sys.argv[4],'w',ZIP_DEFLATED) as output:
    for entry in original.infolist():
        content=original.read(entry.filename)
        if entry.filename=='xl/styles.xml': content=content.replace(b'</styleSheet>',extension+b'</styleSheet>')
        if entry.filename=='xl/workbook.xml': content=content.replace(b'</workbook>',workbook_extension+b'</workbook>')
        output.writestr(entry,content)
mapped=Workbook(); first=mapped.active; first.title='模板页'
first.append(['ID','Amount','Amount','Note']); first.append(['A001',10,20,'first']); first.append(['A002',30,40,'first extra'])
second=mapped.create_sheet('另一页'); second.append(['ID','Amount','Amount','Note']); second.append(['B001',200,100,'second']); second.append(['B002',400,300,'second extra'])
mapped.save(sys.argv[5])
extra=Workbook(); sheet=extra.active; sheet.title='其他来源'
sheet.append(['ID','Amount','Amount','Note']); sheet.append(['C001',50,60,'third']); sheet.append(['C002',70,80,'excluded'])
extra.save(sys.argv[6])
for target,headers,row in [(sys.argv[7],['ID','Task Name','Unit','Amount','Month','State'],['N001','first','DeptA',12,1,'todo']),
                           (sys.argv[8],['ＩＤ','Task\\n  Name','Unit','Amount','Month','State'],['N002','second','DeptB',34,2,'todo'])]:
    book=Workbook(); sheet=book.active; sheet.title='规范化'
    sheet.append(headers); sheet.append(row); book.save(target)
book=Workbook(); sheet=book.active; sheet.title='方向待确认'
sheet.append(['项目','甲','乙','丙'])
for i in range(1,5): sheet.append([f'项{i}','数据','数据','数据'])
book.save(sys.argv[9])
for target,headers,row in [(sys.argv[10],['ID','Plan','Done'],['R001',10,6]),(sys.argv[11],['ID','Target','Actual'],['R002',20,9])]:
    book=Workbook(); sheet=book.active; sheet.title='明细'; sheet.append(headers); sheet.append(row); book.save(target)
`, source, shifted, duplicate, compatibility, sameBook, otherMapped, normalizedA, normalizedB, uncertain, driftTemplate, driftSource], {cwd: root, encoding: 'utf8'});
  const params = {source, shifted, duplicate, compatibility, sameBook, otherMapped, normalizedA, normalizedB, uncertain, driftTemplate, driftSource,
    feedback: ['工程技术部', '财务部', '市场营销部'].map(n => path.join(base, `W1_反馈_${n}.xlsx`)), out};
  const result = await win.webContents.executeJavaScript(`(async (P) => {
    const wait = async (fn, label) => {
      for (let i=0;i<200;i++) { if(fn()) return; await new Promise(r=>setTimeout(r,100)); }
      throw Error('等待超时：'+label);
    };
    const assert=(ok,msg)=>{if(!ok)throw Error(msg)};
    const change=(element,value)=>{element.value=value;element.dispatchEvent(new Event('change',{bubbles:true}));};
    const modalHidden=()=>['done-pane','progress-pane'].every(id=>getComputedStyle(document.getElementById(id)).display==='none');
    await wait(()=>document.querySelector('.tool-card'),'工具入口');
    document.querySelector('.tool-card').click();
    await wait(()=>window.__excelTest,'工作台加载');
    const T=window.__excelTest;
    await T.switchMode('split');
    await T.addPaths({files:[P.source]});
    document.querySelector('#header-range').value='A4:G4';
    document.querySelector('#data-range').value='A5:G16';
    await T.applyRanges();
    T.selectField(3,'rows');
    assert(await T.previewResult(),'拆分预览失败：'+document.querySelector('#notice').textContent);
    assert(T.current().result.can_export,'拆分有阻断：'+JSON.stringify(T.current().result.issues));
    assert(T.current().result.groups.length===5,'拆分应有5个单位');
    document.querySelector('#scheme-name').value='结构复用验收';
    T.saveScheme();
    const schemeId=document.querySelector('#scheme-select').value;
    assert(schemeId,'方案未保存');
    assert(await T.exportResult(P.out+'/split'),'拆分导出失败：'+document.querySelector('#notice').textContent);
    document.querySelector('#btn-again').click();
    assert(modalHidden(),'拆分完成后有浮层残留');

    await T.addPaths({files:[P.shifted]});
    document.querySelector('#scheme-select').value=schemeId;
    assert(await T.applyScheme(),'移位方案复用失败：'+document.querySelector('#notice').textContent);
    assert(T.current().group[0].index===6,'原第3列分组字段未重定位到第6列');
    assert(T.config.selection.header_start===7&&T.config.selection.data_start===8&&T.config.selection.data_end===20,'范围未适应新表头和新增记录：'+JSON.stringify(T.config.selection));
    assert(await T.previewResult(),'移位方案预览失败');
    assert(T.current().result.groups.length===5,'移位后拆分分组不一致');
    assert(T.current().result.groups.reduce((n,g)=>n+g.count,0)===13,'新增加的一条记录未纳入方案范围');
    await T.addPaths({files:[P.duplicate]});
    document.querySelector('#scheme-select').value=schemeId;
    assert(!(await T.applyScheme()),'重名表头被静默套用');
    assert(T.current().schemeProblems.length>0&&document.querySelector('#btn-preview').disabled,'重名表头没有暂停并要求补选');
    await T.addPaths({files:[P.source]});
    document.querySelector('#scheme-select').value=schemeId;
    assert(await T.applyScheme(),'原结构方案恢复失败');

    await T.switchMode('reconcile');
    await T.addPaths({files:[P.source]});
    document.querySelector('#header-range').value='A4:G4';
    document.querySelector('#data-range').value='A5:G16';
    await T.applyRanges();
    for(const f of P.feedback) {
      await T.addPaths({files:[f]},'feedback');
      document.querySelector('#header-range').value='A4:G4';
      document.querySelector('#data-range').value='A5:G'+T.config.report.max_row;
      await T.applyRanges();
    }
    await T.showSource(T.current().files[0]);
    T.current().role='keys'; T.selectField(1,'rows');
    T.current().role='updates'; for(const n of [2,3,4,5,7])T.selectField(n,'rows');
    T.current().role='compares'; for(const n of [2,3,4,5])T.selectField(n,'rows');
    // The optional second output is configured through real controls.
    document.querySelector('#filtered-enabled').click();
    document.querySelector('#add-output-filter').click();
    const filter=document.querySelector('#filtered-rows .filter-row');
    change(filter.querySelectorAll('select')[0],'4');
    change(filter.querySelectorAll('select')[1],'month');
    const filterValue=document.querySelector('#filtered-rows input');filterValue.value='2026-06';filterValue.dispatchEvent(new Event('input',{bubbles:true}));
    assert(await T.previewResult(),'回填预览失败：'+document.querySelector('#notice').textContent);
    let report=T.current().result;
    assert(report.changes.some(c=>c.key[0]==='T002'&&c.after===25&&c.marked),'未识别计划调整');
    assert(report.changes.some(c=>c.key[0]==='T005'&&c.marked===false),'正常状态反馈误标红');
    assert(document.querySelectorAll('.changed-value').length>0,'差异没有红字预览');
    const addition=report.issues.find(i=>i.kind==='addition');
    assert(addition,'新增记录应要求确认');
    const issueSelect=[...document.querySelectorAll('#issues-panel select')].find(select=>select.getAttribute('aria-label')===addition.message);
    assert(issueSelect,'未显示新增记录处理控件');
    change(issueSelect,'append');
    await wait(()=>!T.state.busy&&T.current().result?.can_export,'通过界面确认新增并重新预览');
    assert(T.current().resolutions[addition.id]==='append','界面事件未提交新增处理方式');
    const missing=T.current().result.missing;
    assert(missing.length>0,'验收应包含未反馈记录');
    const missingDetails=document.querySelector('#missing-records details');
    assert(missingDetails,'未反馈记录只有数量，没有具体清单');
    missingDetails.querySelector('summary').click();
    assert(missingDetails.open,'未反馈记录清单无法展开');
    assert(missingDetails.querySelectorAll('tr').length===missing.length+1,'未反馈清单没有列出所有具体记录');
    assert(missingDetails.textContent.includes(missing[0].key[0]),'未反馈清单缺少记录标识');
    const changes=T.current().result.changes.length;
    assert(await T.exportResult(P.out+'/全年更新.xlsx'),'回填导出失败：'+document.querySelector('#notice').textContent);
    assert(document.querySelector('#done-out').textContent.includes('筛选结果'),'完成界面未显示第二份输出');
    // Switching modes must also clear a completed modal without relying on a hook to hide it.
    await T.switchMode('merge');
    assert(modalHidden(),'切换模式后仍有完成或进度浮层');
    await T.addPaths({files:[P.source,P.feedback[1]]});
    document.querySelector('[name="merge-output"][value="single"]').click();
    assert(await T.previewResult(),'合并预览失败：'+document.querySelector('#notice').textContent);
    assert(T.current().result.can_export,'合并有阻断：'+JSON.stringify(T.current().result.issues));
    const mergeReport=T.current().result,mergeSummary=document.querySelector('#result-summary').textContent;
    assert(mergeReport.totals.rows_written>0&&mergeSummary.includes(mergeReport.totals.rows_written+' 条记录'),'合并汇总未显示真实写入记录数：'+mergeSummary);
    const gridSample=mergeReport.groups[0];
    assert(gridSample.sample_kind==='grid','合并预览应按输出工作表坐标显示');
    assert(document.querySelector('#result-samples .hint').textContent.includes('工作表共 '+gridSample.row_count+' 行'),'合并样本未显示实际工作表行数');
    assert(document.querySelector('#result-samples .sample-table tr').textContent.startsWith('行AB'),'合并样本错误套用了输入业务表头');
    assert(document.querySelector('#result-samples .sample-tabs button').textContent.includes('共 '+gridSample.row_count+' 行'),'未知记录数被显示为零');
    assert(await T.exportResult(P.out+'/合并结果.xlsx'),'合并导出失败');
    await T.switchMode('merge');
    document.querySelector('#btn-clear').click();
    await T.addPaths({files:[P.compatibility]});
    document.querySelector('[name="merge-output"][value="workbook"]').click();
    assert(await T.previewResult(),'扩展样式预览失败：'+document.querySelector('#notice').textContent);
    const compatibleReport=T.current().result;
    const compatibilityWarning=compatibleReport.warnings.find(w=>typeof w==='string'&&/扩展样式|styles.xml|extLst/.test(w));
    const workbookWarning=compatibleReport.warnings.find(w=>typeof w==='string'&&w.includes('workbook.xml'));
    assert(compatibilityWarning,'扩展样式必须给出可见兼容提示');
    assert(workbookWarning,'工作簿未知扩展必须明确给出 workbook.xml 提示');
    assert(compatibleReport.can_export&&!document.querySelector('#btn-export').disabled,'纯样式兼容提示仍然阻止导出');
    assert(document.querySelector('#result-warnings').textContent.includes(compatibilityWarning),'预览没有展示样式兼容提示');
    assert(document.querySelector('#result-warnings').textContent.includes(workbookWarning),'预览没有展示 workbook.xml 扩展提示');
    assert(document.querySelector('#result-warnings').textContent.includes('可继续导出'),'兼容提示没有明确允许继续导出');
    assert(getComputedStyle(document.querySelector('#issues-panel')).display==='none','纯兼容提示被误放进阻断问题区');
    assert(await T.exportResult(P.out+'/兼容样式合并结果.xlsx'),'有兼容提示的实际导出失败：'+document.querySelector('#notice').textContent);
    const completedWarnings=document.querySelector('#done-warnings');
    assert(getComputedStyle(completedWarnings).display!=='none'&&completedWarnings.open,'完成页的兼容提示没有直接显示');
    assert(document.querySelector('#done-warnings-list').textContent.includes(compatibilityWarning),'导出完成后丢失兼容提示');
    assert(document.querySelector('#done-warnings-list').textContent.includes(workbookWarning),'导出完成后丢失 workbook.xml 扩展提示');
    completedWarnings.querySelector('summary').click();assert(!completedWarnings.open,'完成页提示不可折叠');
    completedWarnings.querySelector('summary').click();assert(completedWarnings.open,'完成页提示不能重新展开');

    // The template is a worksheet, not the whole file: another sheet needs its own mapping controls.
    await T.switchMode('merge');document.querySelector('#btn-clear').click();
    await T.addPaths({files:[P.sameBook,P.otherMapped]});
    document.querySelector('[name="merge-output"][value="single"]').click();
    const mappedSource=T.current().files[0];
    document.querySelector('#header-range').value='A1:D1';document.querySelector('#data-range').value='A2:D3';
    assert(await T.applyRanges(),'模板范围应用失败');
    assert(document.querySelector('#mapping-details').classList.contains('hidden'),'真正模板页不应显示自我映射');
    await T.showSource(mappedSource,'files',0,'另一页');
    document.querySelector('#header-range').value='A1:D1';document.querySelector('#data-range').value='A2:D3';
    assert(await T.applyRanges(),'同簿另一页范围应用失败');
    const mappingPanel=document.querySelector('#mapping-details');
    assert(!mappingPanel.classList.contains('hidden'),'同一文件的另一页没有字段对应入口');
    mappingPanel.querySelector('summary').click();
    const samePageSelects=[...document.querySelectorAll('#mapping-table select')];
    assert(samePageSelects.length===4,'同簿另一页没有完整的字段对应控件');
    assert(!samePageSelects[1].value&&!samePageSelects[2].value,'重名字段被自动按顺序猜测');
    change(samePageSelects[1],'3');change(samePageSelects[2],'2');
    const savedMapping=JSON.stringify(T.config.field_map);
    const otherSource=T.current().files[1];
    await T.showSource(otherSource,'files',1);
    document.querySelector('#header-range').value='A1:D1';document.querySelector('#data-range').value='A2:D3';
    assert(await T.applyRanges(),'另一来源范围应用失败');
    let otherSelects=[...document.querySelectorAll('#mapping-table select')];
    change(otherSelects[1],'2');change(otherSelects[2],'3');
    const otherMapping=JSON.stringify(T.config.field_map);
    document.querySelector('#data-range').value='A2:D2';document.querySelector('#data-range').dispatchEvent(new Event('input',{bubbles:true}));
    assert(await T.applyRanges(),'另一来源仅修改数据范围失败');
    assert(JSON.stringify(mappedSource.configs['另一页'].field_map)===savedMapping,'修改另一来源范围清掉了已有映射');
    assert(JSON.stringify(T.config.field_map)===otherMapping,'只修改数据行清掉了当前字段映射');
    assert(T.config.selection_origin==='manual','应用范围没有标为手动');
    assert(await T.previewResult(),'同簿手动映射预览失败：'+document.querySelector('#notice').textContent);
    assert(T.current().result.can_export,'同簿手动映射仍被阻断：'+JSON.stringify(T.current().result.issues));
    assert(await T.exportResult(P.out+'/同簿字段对应结果.xlsx'),'同簿字段对应实际导出失败');

    await T.switchMode('merge');document.querySelector('#btn-clear').click();
    await T.addPaths({files:[P.normalizedA,P.normalizedB]});
    assert(T.config.selection_origin==='auto','初次识别被当作手动确认');
    const normalizedSource=T.current().files[1];
    await T.showSource(normalizedSource,'files',1);
    assert(T.config.selection_origin==='auto','切换来源改变了自动范围来源');
    assert([...document.querySelectorAll('#mapping-table select')].every(select=>select.value),'全半角/空白不同字段没有在界面自动对应');
    await T.showSource(normalizedSource,'files',1,normalizedSource.sheet,{row:2,col:1});
    assert(T.config.selection_origin==='auto','分页被当作手动确认');
    const automaticPayload=await T.buildPayload();
    assert(automaticPayload.selections.every(item=>item.selection_origin==='auto'),'默认合并未保留自动识别来源');
    assert(await T.previewResult(),'规范字段默认预览失败：'+document.querySelector('#notice').textContent);
    assert(T.current().result.can_export&&T.current().result.groups.length===1,'规范字段被分组或阻断');
    assert(T.current().result.groups[0].columns.length===6,'规范字段未合并成六列');
    assert(await T.exportResult(P.out+'/规范字段合并结果.xlsx'),'规范字段实际导出失败');

    await T.switchMode('merge');document.querySelector('#btn-clear').click();
    await T.addPaths({files:[P.uncertain]});
    const recognitionNotice=document.querySelector('#source-warnings');
    assert(T.config.report.recognition.orientation_confident===false,'测试输入应保留方向不确定状态');
    assert(getComputedStyle(recognitionNotice).display!=='none'&&recognitionNotice.textContent.includes('方向'),'原表预览没有显示低置信识别提示');
    await T.showSource(T.current().files[0],'files',0,T.source.sheet,{row:2,col:1});
    assert(recognitionNotice.textContent.includes('方向')&&T.config.selection_origin==='auto','翻页丢失识别提醒或冒充确认');
    assert(await T.previewResult(),'待确认方向预览失败');
    assert(T.current().result.can_export&&document.querySelector('#result-warnings').textContent.includes('方向'),'识别提醒丢失或错误阻止默认导出');
    document.querySelector('#header-range').value='A1:D1';document.querySelector('#data-range').value='A2:D5';
    assert(await T.applyRanges(),'手工确认方向范围失败');
    assert(T.config.selection_origin==='manual'&&T.config.report.recognition.confirmed,'手动范围未被明确确认');
    assert(!recognitionNotice.textContent.includes('方向'),'手动确认后仍重复显示已解决的识别提示');
    await T.switchMode('split');
    assert(T.current().group[0].index===3,'模式切换丢失拆分设置');
    assert(modalHidden(),'验收结束时仍有浮层');
    assert(document.querySelector('#grid-wrap td'),'原表网格没有渲染');
    assert(document.querySelector('#grid-wrap td[colspan]'),'原表合并区域没有渲染');
    document.querySelector('#workspace').scrollIntoView({block:'start'});
    await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
    await new Promise(resolve=>setTimeout(resolve,350));
    return {groups:5,changes,missing:missing.length,schemeSaved:!!localStorage.getItem('excelToolSchemes.v1'),shiftedScheme:true,duplicateBlocked:true,domIssueChoice:true,compatibilityWarningExport:true,workbookExtensionWarningExport:true,completionWarningVisible:true,sameWorkbookMapping:true,dataRangePreservesMappings:true,automaticSelectionOrigin:true,normalizedSixColumns:true,recognitionWarningVisible:true,modalsHidden:modalHidden(),out:P.out};
  })(${JSON.stringify(params)})`, true);
  await win.webContents.executeJavaScript(`(async(P)=>{
    const T=window.__excelTest,assert=(ok,message)=>{if(!ok)throw Error(message)};
    const change=(element,value)=>{element.value=value;element.dispatchEvent(new Event('change',{bubbles:true}));};
    await T.switchMode('merge');document.querySelector('#btn-clear').click();
    await T.addPaths({files:[P.driftTemplate,P.driftSource]});
    document.querySelector('[name="merge-output"][value="single"]').click();
    await T.showSource(T.current().files[1],'files',1);
    document.querySelector('#mapping-details').open=true;
    const controls=document.querySelectorAll('#mapping-table select');
    change(controls[1],'2');change(controls[2],'3');
    const payload=await T.buildPayload(),mapped=payload.selections.find(item=>item.path===P.driftSource);
    assert(mapped.field_map_context?.source.length===3&&mapped.field_map_context?.target.length===3,'显式映射缺少双方字段身份快照');
    assert(JSON.parse(mapped.field_map_context.source[1].match_signature)[0]==='target','来源身份快照不对应映射建立时的字段');
    assert(JSON.parse(mapped.field_map_context.target[1].match_signature)[0]==='plan','模板身份快照不对应映射建立时的字段');
    assert(T.config.selection_origin==='auto','设置字段对应不应冒充手动确认范围');
  })(${JSON.stringify(params)})`,true);
  // Modify only this test's generated file, after the renderer has established positional mappings.
  execFileSync(process.env.TOOLBOX_PYTHON || 'python',['-c',`
import sys
from openpyxl import load_workbook
book=load_workbook(sys.argv[1]); sheet=book.active
for row in range(1,sheet.max_row+1):
    before=sheet.cell(row,2).value; sheet.cell(row,2,sheet.cell(row,3).value); sheet.cell(row,3,before)
book.save(sys.argv[1]); book.close()
`,driftSource],{cwd:root,encoding:'utf8'});
  Object.assign(result,await win.webContents.executeJavaScript(`(async(P)=>{
    const T=window.__excelTest,assert=(ok,message)=>{if(!ok)throw Error(message)};
    assert(T.config.report.fields[1].label==='Target'&&T.config.field_map[2]===2,'外部换列前提失败：界面应仍保存旧表头与位置');
    assert(await T.previewResult(),'外部换列后直接预览失败：'+document.querySelector('#notice').textContent);
    assert(T.current().result.can_export,'唯一字段外部换列未正确迁移：'+JSON.stringify(T.current().result.issues));
    const row=T.current().result.groups.flatMap(group=>group.sample).find(values=>values[0]==='R002');
    assert(row&&row[1]===20&&row[2]===9,'外部换列后旧坐标静默错配：'+JSON.stringify(row));
    assert(await T.exportResult(P.out+'/外部换列合并结果.xlsx'),'外部换列实际导出失败：'+document.querySelector('#notice').textContent);
    document.querySelector('#btn-again').click();
    await T.showSource(T.current().files[1],'files',1);
    assert(T.config.field_map[2]===3&&T.config.field_map[3]===2,'重新读取没有按字段身份迁移界面映射');
    const controls=document.querySelectorAll('#mapping-table select');
    assert(controls[1].value==='3'&&controls[2].value==='2','字段对应控件仍显示旧位置');
    assert(!T.current().result&&document.querySelector('#btn-export').disabled,'重新读取结构变化后旧结果未失效');
    assert(await T.applyRanges(),'重新读取后应用范围失败');
    assert(T.config.field_map[2]===3&&T.config.field_map[3]===2,'应用范围造成映射二次迁移');
    assert(await T.previewResult()&&T.current().result.can_export,'迁移后的界面映射未通过新预览');
    const reread=T.current().result.groups.flatMap(group=>group.sample).find(values=>values[0]==='R002');
    assert(reread&&reread[1]===20&&reread[2]===9,'重新读取后的映射改变结果');
    await T.switchMode('split');document.querySelector('#workspace').scrollIntoView({block:'start'});
    await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
    return {mappingContextSent:true,externalColumnChangeHandled:true,rereadMappingRelocated:true};
  })(${JSON.stringify(params)})`,true));
  execFileSync(process.env.TOOLBOX_PYTHON || 'python',['-c',`
import sys
from openpyxl import load_workbook
book=load_workbook(sys.argv[1]); book.active.cell(1,3,'RevisedTarget'); book.save(sys.argv[1]); book.close()
`,driftSource],{cwd:root,encoding:'utf8'});
  Object.assign(result,await win.webContents.executeJavaScript(`(async()=>{
    const T=window.__excelTest,assert=(ok,message)=>{if(!ok)throw Error(message)};
    await T.switchMode('merge');
    assert(T.config.mapping_review&&document.querySelector('#btn-preview').disabled,'无法恢复的显式映射未暂停等待核对');
    assert(getComputedStyle(document.querySelector('#mapping-review')).display!=='none','字段对应核对提示不可见');
    const select=document.querySelectorAll('#mapping-table select')[1];
    select.value='3';select.dispatchEvent(new Event('change',{bubbles:true}));
    document.querySelector('#mapping-reviewed').click();
    assert(!T.config.mapping_review&&!document.querySelector('#btn-preview').disabled,'核对当前映射后仍无法预览');
    assert(await T.previewResult()&&T.current().result.can_export,'重新核对映射后未能恢复正常预览');
    const row=T.current().result.groups.flatMap(group=>group.sample).find(values=>values[0]==='R002');
    assert(row&&row[1]===20&&row[2]===9,'重新核对后的映射值错误');
    await T.switchMode('split');document.querySelector('#workspace').scrollIntoView({block:'start'});
    await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
    return {mappingReviewResolved:true};
  })()`,true));
  // Allow the hidden window's compositor to finish after modal teardown and scrolling.
  await new Promise(resolve=>setTimeout(resolve,500));
  fs.writeFileSync(path.join(out, '工作台.png'), (await win.webContents.capturePage()).toPNG());
  if (errors.length) throw Error(errors.join('\n'));
  const actual = fs.readdirSync(out);
  if (!actual.includes('全年更新.xlsx') || !actual.includes('全年更新_筛选结果.xlsx') ||
      !actual.includes('合并结果.xlsx') || !actual.includes('兼容样式合并结果.xlsx')) throw Error('缺少预期导出文件：'+actual.join(','));
  execFileSync(process.env.TOOLBOX_PYTHON || 'python', ['-c', `
import sys
from openpyxl import load_workbook
book=load_workbook(sys.argv[1],data_only=False)
try:
    sheet=next(ws for ws in book if ws['A2'].value=='X001')
    assert list(sheet.values)==[('编号','分类','数值','计算'),('X001','甲',12,'=C2*2'),('X002','乙',34,'=SUM(C2:C3)')], 'Exported values or formulas changed'
finally:
    book.close()
book=load_workbook(sys.argv[2],data_only=False)
try:
    sheet=book.active
    assert sheet.max_column==6 and sheet.max_row==3, 'Normalized headings did not yield six columns and two records'
    assert list(sheet.values)[1:]==[('N001','first','DeptA',12,1,'todo'),('N002','second','DeptB',34,2,'todo')], 'Normalized field alignment changed values'
finally:
    book.close()
book=load_workbook(sys.argv[3],data_only=False)
try:
    sheet=book.active; records={row[0]:row[1:] for row in list(sheet.values)[1:]}
    assert records['B001']==(100,200,'second') and records['B002']==(300,400,'second extra'), 'Same-workbook positional mapping was not applied'
    assert records['C001']==(50,60,'third') and 'C002' not in records, 'Edited data range was not applied'
finally:
    book.close()
book=load_workbook(sys.argv[4],data_only=False)
try:
    records={row[0]:row[1:] for row in list(book.active.values)[1:]}
    assert records['R002']==(20,9), 'External column reorder changed mapped values'
finally:
    book.close()
`, path.join(out, '兼容样式合并结果.xlsx'), path.join(out, '规范字段合并结果.xlsx'), path.join(out, '同簿字段对应结果.xlsx'), path.join(out, '外部换列合并结果.xlsx')], {cwd: root, encoding: 'utf8'});
  result.exportedValuesAndFormulasVerified=true;
  fs.writeFileSync(path.join(out, 'result.json'), JSON.stringify(result,null,2), 'utf8');
  return result;
};
