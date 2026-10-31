// Fictional source documents only. No working spreadsheet IDs or personnel records.
module.exports = () => [
  {id:'SYNTHETIC_MAIN_0123456789',name:'Вымышленный резерв',sheets:[
    {id:0,name:'Участники',rows:[
      ['Кадровый резерв — пример'],[],
      ['ФИО','Регион','Баллы','Дата включения','Активен','Телефон','','Регион','Итог'],
      ['Участник 01','Астана',50,new Date('2026-01-15T00:00:00Z'),true,{value:123,display:'00123'},'Заметка','Север',{value:90,formula:'=45*2'}],
      ['Участник 02','Алматы',0,new Date('2026-02-15T00:00:00Z'),false,{value:456,display:'00456'},'','Юг',80],
      [],['Участник 03','Астана',100,'','',789,'Текст','Север',75]
    ]},
    {id:42,name:'Иные поля',rows:[['Проект','Статус'],['Пример проекта','В работе']]},
    {id:99,name:'Пустой',rows:[]}
  ]},
  {id:'SYNTHETIC_OTHER_0123456789',name:'Другой источник',sheets:[{id:0,name:'Реестр',rows:[['Name','Skill'],['Example Person','Research']]}]}
];
