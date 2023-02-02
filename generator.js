const xl = require('excel4node');
const fs = require('fs');

function loadData() {
  return JSON.parse(fs.readFileSync('sources.json'));
}
const data = loadData();
const wb = new xl.Workbook();
const ws = wb.addWorksheet("Информационные источники");

const headColumnsName =
{
  sectionName: {
    indexCell: 1,
    label: 'Раздел',
    width: 25,
    id: 'sectionName'
  },
  nameSkill: {
    indexCell: 2,
    label: 'Навык',
    width: 25,
    id: 'skillName'
  },
  infoLinks: {
    indexCell: 3,
    label: 'Информационные ссылки',
    width: 25,
    id: 'skillName'
  }
};
const style = wb.createStyle({
  alignment: { horizontal: "center", vertical: 'center', wrapText: true },
});
ws.column(1).setWidth(25);
ws.column(2).setWidth(50);
ws.column(3).setWidth(100);

Object.keys(headColumnsName).forEach((item, index) => {
  ws.cell(1, index + 1).string(headColumnsName[item].label).style(style);
});

let sectionColumnIndex = 2;
let skillColumnIndex = 2;
let linkColumnIndex = 2;

const getRowMerge = (item, key) => {
  return item.sectionSkills.reduce((acc, item) => {
    acc += item[key].length;
    return acc;
  }, 0);
};

data.forEach(item => {
  const rowMerged = getRowMerge(item, 'infoLinks');
  const curIdx = sectionColumnIndex++;
  ws.cell(curIdx, 1, curIdx + rowMerged - 1, 1, true).string(item.sectionName).style(style);
  item.sectionSkills.forEach((skill) => {
    const curSkillIdx = skillColumnIndex++;
    const secondCell = skillColumnIndex + skill.infoLinks.length - 1;
    if (curSkillIdx === secondCell) {
      ws.cell(curSkillIdx, 2).string(skill.nameSkill).style(style);
    } else {
      ws.cell(curSkillIdx, 2, curSkillIdx + skill.infoLinks.length - 1, 2, true).string(skill.nameSkill).style(style);
    }
    skill.infoLinks.forEach((link) => {
      const curLinkIdx = linkColumnIndex++
      ws.cell(curLinkIdx, 3).link(link).style(style);
    });
    skillColumnIndex += skill.infoLinks.length - 1;
  });
  sectionColumnIndex += rowMerged - 1;
});
wb.write('sources.xlsx');
