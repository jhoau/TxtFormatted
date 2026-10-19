// Prueba de interacción de la lógica de la interfaz sin dependencias externas.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

class Element {
  constructor(tag) { this.tag = tag; this.children = []; this.listeners = {}; this.scrollTop = 0; this.classList = {toggle(){}, add(){}, remove(){}}; }
  append(...items) { this.children.push(...items); }
  replaceChildren(...items) { this.children = items; }
  addEventListener(event, action) { this.listeners[event] = action; }
  setAttribute(key, value) { this[key] = value; }
  focus() {}
  remove() {}
  async click() { if (!this.disabled) await this.listeners.click?.({target:this}); }
}
const elements = new Map();
const get = id => { if(!elements.has(id)) elements.set(id,new Element('div')); return elements.get(id); };
const original = {columns:['Nombre','Email','Teléfono'], rows:[{'Nombre':'Ángel','Email':'a@example.com','Teléfono':'00123'}], warnings:[], encoding:'UTF-8', note:''};
const exportsSent = [];
const context = vm.createContext({
  document:{getElementById:get,createElement:tag=>new Element(tag),body:new Element('body')},
  File, Set, URL:{createObjectURL:()=> 'blob:test',revokeObjectURL(){}},setTimeout(){},
  fetch:async (path,options)=>{
    if(path==='/api/parse')return {ok:true,json:async()=>structuredClone(original)};
    exportsSent.push(JSON.parse(options.body));
    return {ok:true,blob:async()=>new Blob(['test'])};
  },
});
const html = fs.readFileSync('index.html','utf8');
vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],context);
const rows = () => get('columns').children;
const checkbox = index => rows()[index].children[0].children[0];
const move = (index,direction) => rows()[index].children[1].children[direction<0?0:1].click();
const headers = () => get('preview').children[0].children[0].children[0].children.map(cell=>cell.textContent);

(async()=>{
  await get('example').click();
  assert.deepEqual(headers(),['#','Nombre','Email','Teléfono']);
  assert.equal(rows()[0].children[1].children[0].disabled,true);
  assert.equal(rows()[2].children[1].children[1].disabled,true);
  await move(2,-1);
  assert.deepEqual(headers(),['#','Nombre','Teléfono','Email']);
  await move(1,-1);
  assert.deepEqual(headers(),['#','Teléfono','Nombre','Email']);
  // Desmarcar y mover una columna conserva sus valores y el orden de las otras.
  checkbox(1).checked=false;checkbox(1).listeners.change();
  await move(1,1);
  assert.deepEqual(headers(),['#','Teléfono','Email']);
  await get('export-csv').click();
  await get('export').click();
  for(const payload of exportsSent){
    assert.deepEqual(payload.columns,['Teléfono','Email']);
    assert.equal(payload.rows[0]['Teléfono'],'00123');
    assert.equal(payload.rows[0]['Nombre'],'Ángel');
  }
  assert.deepEqual(exportsSent.map(p=>p.format),['csv','xlsx']);
  await get('all').click();
  assert.deepEqual(headers(),['#','Teléfono','Email','Nombre']);
  await get('none').click();
  assert.equal(get('export-csv').disabled,true);
  assert.equal(get('export').disabled,true);
  await get('all').click();
  assert.deepEqual(headers(),['#','Teléfono','Email','Nombre']);
  await get('example').click();
  assert.deepEqual(headers(),['#','Nombre','Email','Teléfono']);
  console.log('Interfaz: mover columnas, límites, selección, orden CSV/Excel y carga de otro archivo correctos.');
})().catch(error=>{console.error(error);process.exitCode=1;});
