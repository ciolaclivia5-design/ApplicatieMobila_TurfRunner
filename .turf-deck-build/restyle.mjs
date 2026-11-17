import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const workspaceDir = process.cwd();
const build = path.join(workspaceDir, '.turf-deck-build', 'restyle');
const output = path.join(workspaceDir, 'presentation-output', 'Turf Runner - Week 8 - Canva Style.pptx');
const runtime = '/Users/livia-maria/.cache/codex-runtimes/codex-primary-runtime/dependencies';
const tools = '/Users/livia-maria/.codex/plugins/cache/openai-primary-runtime/presentations/26.1116.12020/skills/presentations/container_tools';

const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const navy = '#050F21', panel = '#091B31', white = '#F5F8FC', soft = '#C6D4E5';
const cyan = '#00B5D7', lightCyan = '#45D8EE', edge = '#087C99', font = 'Arial';

const originalScript = await fs.readFile(path.join(build, '..', 'build.mjs'), 'utf8');
const notes = [...originalScript.matchAll(/^(?:let )?s = make\((?:null|'[^']*'), '(.*)'(?:, true)?\);$/gm)]
  .map(match => match[1].replaceAll('\\n', '\n'));
if (notes.length !== 10) throw new Error('Expected ten sets of speaker notes');

function text(slide, value, x, y, w, h, o = {}) {
  const shape = slide.shapes.add({
    geometry: 'textbox', name: o.name ?? value.split('\n')[0],
    position: { left: x, top: y, width: w, height: h }, fill: 'none',
    line: { fill: 'none', width: 0 },
  });
  shape.text = value;
  shape.text.style = {
    typeface: font, fontSize: o.size ?? 25, bold: o.bold ?? false,
    color: o.color ?? white, alignment: o.align ?? 'left', insets: 0, verticalAlignment: 'top',
  };
  return shape;
}
function rect(slide, name, x, y, width, height, fill, border = 'none', borderWidth = 0, radius = 0) {
  return slide.shapes.add({
    geometry: radius ? 'roundRect' : 'rect', name,
    position: { left: x, top: y, width, height }, fill,
    borderRadius: radius || undefined, line: { style: 'solid', fill: border, width: borderWidth },
  });
}
function cut(slide, name, points, fill) {
  slide.shapes.add({ geometry: 'custom', name, position: { left: 0, top: 0, width: 1280, height: 720 },
    fill, line: { width: 0, fill: 'none' }, customPaths: [{
      width: 1280, height: 720,
      commands: [
        { moveTo: { x: points[0][0], y: points[0][1] } },
        ...points.slice(1).map(([x,y]) => ({ lineTo: { x,y } })),
        { close: {} },
      ],
    }],
  });
}
const imgNames = ['map','leaderboard','friends','chat','profile','achievements','notifications'];
const images = Object.fromEntries(await Promise.all(imgNames.map(async name => [
  name, await fs.readFile(path.join(workspaceDir, 'presentation-assets', name + '.png')),
])));
function picture(slide, blob, alt, x, y, width, height, fit = 'cover', rounded = false) {
  slide.images.add({ blob, contentType: 'image/png', alt, fit,
    position: { left:x, top:y, width, height },
    ...(rounded ? { geometry:'roundRect', borderRadius: 16 } : {}),
  });
}
function phone(slide, name, caption, x, y, height) {
  const width = height * 1170 / 2532;
  rect(slide, caption + ' frame', x-6, y-6, width+12, height+12, '#050A15', '#2B4760', 1.6, 21);
  picture(slide, images[name], caption + ' in Turf Runner as User1', x, y, width, height, 'contain', true);
  text(slide, caption, x-20, y+height+14, width+40, 23, { size: 17, color: soft, align:'center' });
}
function make(note) {
  const slide = deck.slides.add();
  slide.background.fill = navy;
  slide.speakerNotes.text = note;
  return slide;
}
function heading(slide, first, second = '', options = {}) {
  const { x = 58, y = 38, w = 1100, size = 54 } = options;
  text(slide, first, x, y, w, size+9, { size, bold:true });
  if (second) text(slide, second, x, y+size-1, w, size+10, { size, bold:true, color:cyan });
}
function section(slide, label, body, x, y, width, bodyHeight = 61) {
  text(slide, label, x, y, width, 29, { size:22, bold:true, color:lightCyan });
  text(slide, body, x, y+32, width, bodyHeight, { size:23, color:soft });
}
function table(slide, values, x, y, widths, rowHeight, fontSize) {
  const width = widths.reduce((a,b)=>a+b,0);
  const t = slide.tables.add({ rows: values.length, columns: widths.length, left: x, top: y,
    width, height: values.length*rowHeight, columnWidths: widths, values });
  t.borders.assign({ fill: edge, style:'solid', width:0.8 });
  values.forEach((_,r)=>{
    t.rows[r].height = rowHeight;
    t.cells.block({ row:r, column:0, rowCount:1, columnCount: widths.length }).assign({
      fill:r === 0 ? '#008CAA' : r%2 ? '#08162A' : '#0A1C32',
      textStyle: { typeface:font, fontSize, bold:r === 0, color:white },
      margins:{left:12,right:10,top:6,bottom:6}, anchor:'middle',
    });
  });
}

let slide = make(notes[0] + '\nCover art: AI-generated decorative route illustration.');
picture(slide, await fs.readFile(path.join(build,'cover.png')), 'Illustrated glowing route over a dark street grid',
  0,0,1280,720);
heading(slide,'Turf','Runner',{x:61,y:143,w:490,size:114});
text(slide,'App Development Progress  /  Week 8', 64, 387, 555, 35, { size:25, bold:true });
text(slide,'A walking and running game\nset in Timișoara',64,466,525,86,{size:30,color:soft});
text(slide,'Livia Maria',64,584,470,38,{size:27,bold:true,color:lightCyan});

slide = make(notes[1] + '\nPhoto: AI-generated illustration of a runner checking a sports watch.');
picture(slide, await fs.readFile(path.join(build,'scope.png')), 'Illustrative photo of a runner checking a sports watch',
  665,0,615,720);
cut(slide,'Cyan angled divider',[[0,0],[842,0],[713,373],[827,720],[0,720]],cyan);
cut(slide,'Navy text field',[[0,0],[827,0],[698,373],[812,720],[0,720]],navy);
heading(slide,'Application','Scope',{x:58,y:56,w:650,size:57});
section(slide,'Track activity','Create an account, start a GPS session and view estimated steps and distance.',60,211,563);
section(slide,'Compete for neighborhoods','Capture Timișoara turfs and compare records on the leaderboard.',60,326,571);
section(slide,'Play with others','Manage friends, join neighborhood chats and check achievements and alerts.',60,440,580);
text(slide,'Tested on the iOS simulator',60,582,520,35,{size:23,bold:true,color:lightCyan});

slide = make(notes[2]);
heading(slide,'System Architecture','and Design');
const architecture = [
  ['React Native app','iOS simulator\nGPS and neighborhood polygons'],
  ['Express REST API','Local Node.js server\nTypeScript routes'],
  ['PostgreSQL','Hosted on Render\nSessions and turf ownership'],
];
architecture.forEach(([label,detail],i)=>{
  const x=60+i*429;
  rect(slide,label,x,224,300,181,panel,edge,1.5,10);
  text(slide,label,x+12,252,276,42,{size:28,bold:true,color:lightCyan,align:'center'});
  text(slide,detail,x+12,311,276,75,{size:22,color:soft,align:'center'});
});
for (const [x,label] of [[360,'HTTP REST'],[789,'TLS']]) {
  text(slide,label,x,269,129,31,{size:17,bold:true,color:soft,align:'center'});
  text(slide,'⟶',x,300,129,45,{size:35,bold:true,color:cyan,align:'center'});
}
for (const [i,[label,detail]] of [
  ['JWT authentication','Authorizes signed-in requests'],
  ['bcrypt hashing','Stores password hashes'],
  ['Polling','Refreshes the map and chat'],
].entries()){
  const x=60+i*396;
  rect(slide,label+' detail',x,480,370,102,'#0A243A',edge,1.25,5);
  text(slide,label,x+12,495,346,30,{size:22,bold:true,color:lightCyan,align:'center'});
  text(slide,detail,x+12,535,346,29,{size:19,color:soft,align:'center'});
}

slide = make(notes[3]);
heading(slide,'Libraries and','Utilities');
table(slide,[
  ['Component','Version','Purpose'],
  ['React Native + React','0.82.1 + 19.1.1','Mobile interface'],
  ['React Navigation','7.x','Tab and stack navigation'],
  ['react-native-maps','1.26.18','Map, neighborhood boundaries and routes'],
  ['Community Geolocation + geolib','3.4.0 + 3.3.4','GPS and point-in-polygon checks'],
  ['AsyncStorage','2.2.0','Local sign-in state'],
  ['Express + TypeScript','4.18.2 + 5.9.3','REST API and server routes'],
  ['pg + PostgreSQL on Render','pg 8.16.3','Database connection and persistence'],
  ['bcrypt + jsonwebtoken','5.1.1 + 9.0.2','Password hashing and authentication tokens'],
  ['Xcode, iOS Simulator, Metro and Git','—','Build, run, test and track changes'],
], 60,171,[405,205,550],47,20);
text(slide,'Versions shown are declared in package.json.',60,656,1000,25,{size:17,color:soft});

slide = make(notes[4].replace('map.png and leaderboard.png','map.png'));
cut(slide,'Navy map field',[[799,0],[1280,0],[1280,720],[597,720]],'#0A1B31');
cut(slide,'Cyan angled divider',[[791,0],[801,0],[606,720],[596,720]],cyan);
heading(slide,'Tracking and','Turf Capture',{x:58,y:71,w:680,size:58});
rect(slide,'Heading rule',60,210,112,3,cyan);
section(slide,'Record a session','The app records GPS points between Start Tracking and Stop Tracking.',60,254,490);
section(slide,'Estimate steps','Distance between points is converted to steps using a 0.7 m stride.',60,371,492);
section(slide,'Capture a turf','Beat the neighborhood’s step record within a session to become its owner.',60,488,490);
phone(slide,'map','Timișoara map',888,42,608);

slide = make(notes[5]);
heading(slide,'Social Features');
text(slide,'How players connect in Turf Runner',60,107,695,40,{size:25,color:soft});
rect(slide,'Social section rule',60,169,110,3,cyan);
section(slide,'Friends','Search by username, respond to requests and view friends’ activity.',60,225,460);
section(slide,'Neighborhood chat','The current GPS location determines which turf’s conversation opens.',60,366,460);
phone(slide,'friends','Friends and requests',597,164,484);
phone(slide,'chat','Iosefin chat',906,164,484);

slide = make('0:45. Four seeded demo accounts, User2 to User5, produced six synthetic runs across five neighborhoods and twelve chat messages. Their runs exercise the capture and takeover rules without affecting User1’s personal stats. Fabric is the current seeded leader at 10,036 steps. The editable table and the screenshot show the same snapshot.\nScreenshot: User1 on the iPhone 17e simulator, leaderboard.png.');
heading(slide,'Progress and','Demo Data');
text(slide,'Seeded demonstration snapshot',60,154,820,38,{size:24,color:soft});
const metrics = [['4','Demo accounts'],['6','Simulated runs'],['5','Owned turfs'],['12','Chat messages']];
metrics.forEach(([num,label],i)=>{
  const y=243+i*86;
  rect(slide,label+' band',60,y,251,70,'#0B2339');
  text(slide,num,72,y+7,69,55,{size:43,bold:true,color:lightCyan});
  text(slide,label,143,y+20,158,42,{size:19,color:white});
});
table(slide,[
  ['Turf','Owner','Record'],
  ['Fabric','User4','10,036'],
  ['Centru','User3','3,474'],
  ['Iosefin','User5','2,895'],
  ['Elisabetin','User2','2,316'],
  ['Complex Studențesc','User3','1,737'],
], 340,245,[254,99,103],53,18);
phone(slide,'leaderboard','Turf leaderboard',935,138,508);
text(slide,'Records are estimated steps.',340,581,440,30,{size:17,color:soft});

slide = make(notes[6] + '\n' + notes[7].replace(/^0:45\. /,'').replace(/^I used synthetic runs.*?chat messages\. /,''));
heading(slide,'Notifications and','Achievements',{size:50});
rect(slide,'Heading rule',60,156,100,3,cyan);
section(slide,'Profile','Distance, steps, captured turfs and sessions for the signed-in player.',60,210,360);
section(slide,'Achievements','Trophy goals and personal progress, including a player’s first capture.',60,351,360);
section(slide,'Notifications','Review game and social alerts. Mark them read one at a time or together.',60,492,360);
phone(slide,'profile','User1 profile',456,181,467);
phone(slide,'achievements','Achievement goals',718,181,467);
phone(slide,'notifications','User1 notifications',980,181,467);

slide = make(notes[8]);
heading(slide,'Challenges and','Issues Encountered');
rect(slide,'Resolved heading and content',60,202,569,429,panel,edge,1.5,6);
rect(slide,'Remaining heading and content',654,202,566,429,panel,edge,1.5,6);
text(slide,'RESOLVED',81,225,500,33,{size:24,bold:true,color:lightCyan});
text(slide,'REMAINING',675,225,500,33,{size:24,bold:true,color:lightCyan});
rect(slide,'Resolved rule',82,272,522,2,edge);
rect(slide,'Remaining rule',675,272,520,2,edge);
section(slide,'Database connectivity','Initialized a replacement Render database and verified its external TLS connection.',82,296,512);
section(slide,'Simulator performance','Reduced load by running one simulator on the MacBook Neo.',82,407,512);
section(slide,'Notification layout','Replaced the banner with a bell and a separate screen.',82,513,512);
section(slide,'Automated tests','TypeScript checks pass. Jest still needs configuration for React Navigation.',675,296,499,94);
section(slide,'Device validation','Physical iPhone GPS and Android checks are still pending.',675,438,499,94);

slide = make(notes[9]);
heading(slide,'Remaining Work &','Schedule');
text(slide,'Proposed schedule',60,157,580,34,{size:23,color:soft});
rect(slide,'Timeline',103,267,3,293,edge);
[
  ['Week 9','Field testing','Test on a physical iPhone and check GPS accuracy and turf boundaries.'],
  ['Week 10','Automated checks','Fix Jest and test authentication, sessions and notifications.'],
  ['Week 11','Submission preparation','Check Android, configure credentials and backups, then rehearse the demo.'],
].forEach(([week,label,detail],i)=>{
  const y=231+i*133;
  slide.shapes.add({geometry:'ellipse',name:week+' milestone',
    position:{left:87,top:y+21,width:36,height:36},fill:navy,line:{style:'solid',fill:cyan,width:2.6}});
  text(slide,week,153,y+22,204,38,{size:29,bold:true,color:lightCyan});
  text(slide,label,365,y+3,812,37,{size:28,bold:true});
  text(slide,detail,365,y+48,804,61,{size:23,color:soft});
});

for(const [i,page] of deck.slides.items.entries()){
  const png = await deck.export({slide:page,format:'png',scale:1});
  await fs.writeFile(path.join(build,'slide-'+String(i+1).padStart(2,'0')+'.png'),new Uint8Array(await png.arrayBuffer()));
}
const candidate = path.join(build,'styled-candidate.pptx');
await (await PresentationFile.exportPptx(deck)).save(candidate);
const { finalizePresentation } = await import(pathToFileURL(path.join(tools,'artifact_tool_utils.mjs')).href);
await finalizePresentation({
  explicitTotalSlideCount:10,workspaceDir,candidatePath:candidate,finalPath:output,
  pythonExecutable:path.join(runtime,'python/bin/python3'),requiredNativeTableOwnerSlides:[4,7],
  integrityValidatorPath:path.join(tools,'inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(tools,'inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-heading-fit','--require-native-table-slide','4','--require-native-table-slide','7'],
  fontPolicy:{basis:'design',families:[font]},verifyArtifactToolImport:true,
  receiptPath:path.join(build,'validation.json'),
});
console.log(output);
