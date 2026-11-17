import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const root = process.cwd();
const build = path.join(root, '.turf-deck-build');
const out = path.join(root, 'presentation-output', 'Turf Runner - Week 8.pptx');
const skill = '/Users/livia-maria/.codex/plugins/cache/openai-primary-runtime/presentations/26.1116.12020/skills/presentations';
const toolDir = path.join(skill, 'container_tools');
const pythonExecutable = '/Users/livia-maria/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3';
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const font = 'Arial';
const night = '#0B1429', white = '#F6FAFF', soft = '#B6C6DB', cyan = '#5EE5F6';
const paper = '#F4F8FC', ink = '#102342', steel = '#425671', teal = '#007B91';

function text(slide, value, x, y, w, h, o = {}) {
  const shape = slide.shapes.add({
    geometry: 'textbox', name: value.split('\n')[0],
    position: { left: x, top: y, width: w, height: h }, fill: 'none',
    line: { fill: 'none', width: 0 },
  });
  shape.text = value;
  shape.text.style = {
    typeface: font, fontSize: o.size ?? 26, bold: o.bold ?? false,
    color: o.color ?? white, alignment: o.align ?? 'left', insets: 0, verticalAlignment: 'top',
  };
  return shape;
}

function make(title, notes, light = false) {
  const slide = deck.slides.add();
  slide.background.fill = light ? paper : night;
  if (title) text(slide, title, 70, 59, 1140, 62, { size: 46, bold: true, color: light ? ink : white });
  text(slide, String(deck.slides.items.length).padStart(2, '0'), 1160, 669, 50, 24,
    { size: 16, color: light ? steel : soft, align: 'right' });
  slide.speakerNotes.text = notes;
  return slide;
}

function sections(slide, entries, x, y, w, step, light = false) {
  entries.forEach(([label, body], i) => {
    const top = y + i * step;
    text(slide, label, x, top, w, 30, { size: 24, bold: true, color: light ? teal : cyan });
    text(slide, body, x, top + 34, w, step - 38, { size: 24, color: light ? steel : soft });
  });
}

const imageFiles = ['map', 'leaderboard', 'chat', 'friends', 'profile', 'achievements', 'notifications'];
const images = Object.fromEntries(await Promise.all(imageFiles.map(async name => [
  name, await fs.readFile(path.join(root, 'presentation-assets', name + '.png')),
])));
function phone(slide, name, label, x, y, height = 489) {
  const width = height * 1170 / 2532;
  slide.images.add({
    blob: images[name], contentType: 'image/png', alt: label + ' in Turf Runner as User1',
    fit: 'contain', geometry: 'roundRect', borderRadius: 12,
    position: { left: x, top: y, width, height },
  });
  text(slide, label, x - 16, y + height + 11, width + 32, 24,
    { size: 17, color: soft, align: 'center' });
}

let s = make(null, '0:35. I’m Livia Maria, and this is my Week 8 progress update for Turf Runner. I built it around a simple idea: a walk or run through Timișoara can change who owns a neighborhood in the game. I’ll show the current simulator build, explain how a GPS session turns into a turf record, and close with the tests and fixes I plan to complete next.');
text(s, 'APP DEVELOPMENT PROGRESS  /  WEEK 8', 76, 133, 1100, 32, { size: 20, bold: true, color: cyan });
text(s, 'Turf Runner', 68, 206, 1140, 136, { size: 116, bold: true });
text(s, 'A walking and running game set in Timișoara', 76, 369, 1110, 50, { size: 35, color: soft });
text(s, 'Livia Maria', 76, 535, 550, 47, { size: 30, bold: true });

s = make('Application scope', '0:40. Users register and sign in, then start a GPS session on the map. The map shows neighborhood boundaries and ownership colors. A runner who beats a neighborhood step record can claim that turf. The leaderboard shows the current owner and record, and the profile summarizes personal activity. Friends, neighborhood chat, trophies and notifications support the competitive loop. The current evidence is from the iOS simulator in a single supported city.');
sections(s, [
  ['Track activity', 'Start and stop GPS sessions. View estimated steps and distance.'],
  ['Compete for neighborhoods', 'Capture turfs in Timișoara and compare records on the leaderboard.'],
  ['Play with others', 'Manage friends, join local chats and view trophies and notifications.'],
], 72, 158, 1000, 131);
text(s, 'Week 8 prototype tested on the iOS simulator', 72, 584, 1100, 34, { size: 21, bold: true, color: cyan });

s = make('System architecture', '0:50. On the left is the React Native client. It gets location updates and includes the Timișoara neighborhood polygons. The client uses an HTTP REST API running locally in this development setup. Express handles authentication, sessions, ownership, and social data. It connects over TLS to PostgreSQL on Render. Passwords are hashed with bcrypt, and authenticated requests use JWTs. The client polls for map and neighborhood chat updates. The client-to-API and API-to-database connections have different transports.', true);
const architecture = [
  ['React Native app', 'iOS simulator\nGPS and map polygons'],
  ['Express API', 'Local Node.js server\nTypeScript routes'],
  ['PostgreSQL', 'Hosted on Render\nSessions and turf ownership'],
];
architecture.forEach(([name, desc], i) => {
  const x = 72 + 400 * i;
  text(s, name, x, 242, 332, 45, { size: 33, bold: true, color: ink });
  text(s, desc, x, 302, 328, 84, { size: 25, color: steel });
});
for (const [x, label] of [[392, 'HTTP REST'], [804, 'TLS']]) {
  text(s, label, x-35, 165, 123, 28, { size: 17, bold: true, color: teal, align: 'center' });
  text(s, '→', x+4, 193, 45, 38, { size: 30, color: teal, align: 'center' });
}
text(s, 'Authentication uses JWTs and bcrypt password hashing. The map and chat refresh by polling.', 72, 531, 1110, 67, { size: 24, color: steel });

s = make('Libraries and utilities', '0:40. React Native renders the interface, and React Navigation controls tabs and stacks. The map and location packages provide the building blocks for turf boundaries and GPS updates. AsyncStorage holds local sign-in state. On the server, Express defines the routes and the pg driver communicates with PostgreSQL. Xcode and the iOS simulator support the demo, Metro serves the development bundle, and Git tracks code changes.', true);
const rows = [
  ['Library or utility', 'Role in Turf Runner'],
  ['React Native + React', 'Mobile interface'],
  ['React Navigation', 'Tab and stack navigation'],
  ['react-native-maps', 'Map, neighborhood polygons and route display'],
  ['Community Geolocation + geolib', 'GPS and point-in-polygon checks'],
  ['AsyncStorage', 'Local sign-in state'],
  ['Express + TypeScript', 'REST API and server routes'],
  ['pg + PostgreSQL on Render', 'Database connection and persistence'],
  ['bcrypt + jsonwebtoken', 'Password hashing and authentication tokens'],
  ['Xcode, iOS Simulator, Metro and Git', 'Build, run, test and track changes'],
];
const table = s.tables.add({ rows: rows.length, columns: 2, left: 72, top: 141,
  width: 1136, height: 470, columnWidths: [458, 678], values: rows });
table.borders.assign({ fill: '#D6E3EE', width: 1, style: 'solid' });
for (let r = 0; r < rows.length; r++) {
  table.rows[r].height = 47;
  const row = table.cells.block({ row: r, column: 0, rowCount: 1, columnCount: 2 });
  row.assign({ fill: r === 0 ? '#15324F' : r % 2 ? '#FFFFFF' : '#EDF4FA',
    textStyle: { typeface: font, fontSize: 22, bold: r === 0, color: r === 0 ? white : ink },
    margins: { left: 13, right: 13, top: 7, bottom: 7 }, anchor: 'middle',
  });
}

s = make('GPS tracking and turf capture', '0:55. On the map, Start Tracking opens an activity session. Location updates record points and let the server estimate distance and steps using a 0.7-meter stride. A point is matched to a neighborhood polygon, and the session accumulates estimated steps in that neighborhood. When the session beats the stored record, the neighborhood changes hands. The leaderboard shows Fabric at 10,036 steps for User4 in the seeded demo.\nScreenshots: User1 on the iPhone 17e simulator, map.png and leaderboard.png.');
sections(s, [
  ['Record a session', 'The app records GPS points between Start Tracking and Stop Tracking.'],
  ['Estimate steps', 'Distance between points is converted to steps using a 0.7 m stride.'],
  ['Capture a turf', 'Beat the neighborhood’s step record within a session to become its owner.'],
], 72, 163, 490, 139);
phone(s, 'map', 'Timișoara map', 641, 142);
phone(s, 'leaderboard', 'Turf leaderboard', 936, 142);

s = make('Friends and neighborhood chat', '0:45. User1 has two accepted friends, User2 and User3, and an incoming request from User4. The Friends screen also supports username search and lets players inspect a friend’s activity. Chat is tied to the player’s GPS neighborhood. Here the simulator is positioned in Iosefin, so User1 sees that neighborhood’s demo messages. The client checks periodically for new messages.\nScreenshots: User1 on the iPhone 17e simulator, friends.png and chat.png.');
sections(s, [
  ['Friends', 'Search by username, respond to requests and view friends’ activity.'],
  ['Neighborhood chat', 'The current GPS location determines which turf’s conversation opens.'],
], 72, 186, 480, 171);
phone(s, 'friends', 'Friends and requests', 641, 142);
phone(s, 'chat', 'Iosefin chat', 936, 142);

s = make('Profile and achievements', '0:45. The profile summarizes the signed-in user’s sessions, distance, steps and owned turfs. The Achievements screen displays goals and progress, including First Capture, Big Capture, Marathon Session, Marathon Runner and Turf Tycoon. The screenshots are captured while signed in as User1. The four seeded runners account for the demo activity, so User1’s personal stats and trophies are at zero.\nScreenshots: User1 on the iPhone 17e simulator, profile.png and achievements.png.');
sections(s, [
  ['Personal statistics', 'Profiles show distance, estimated steps, turfs and sessions.'],
  ['Trophies', 'Players can see goals and progress, including their first capture and a 10,000-step session.'],
], 72, 186, 490, 171);
phone(s, 'profile', 'User1 profile', 641, 142);
phone(s, 'achievements', 'Achievement goals', 936, 142);

s = make('Week 8 demo and notifications', '0:45. I used synthetic runs to verify that location tracking, turf ownership and the leaderboard worked together. Four demo accounts generated six runs, five owned turfs and twelve chat messages. User1 can open the bell on the map to review notifications. The captured screen shows friend-request alerts. Individual alerts can be marked read, and there is also a Mark All Read action.\nScreenshot: User1 on the iPhone 17e simulator, notifications.png.');
const metrics = [['4', 'Demo accounts'], ['6', 'Simulated runs'], ['5', 'Owned turfs'], ['12', 'Chat messages']];
metrics.forEach(([count, label], i) => {
  const x = 72 + (i % 2) * 330, y = 159 + Math.floor(i / 2) * 136;
  text(s, count, x, y, 145, 72, { size: 66, bold: true, color: cyan });
  text(s, label, x+1, y+72, 300, 34, { size: 24, bold: true });
});
text(s, 'The map bell opens the notification history.\nUsers can mark alerts read individually or all at once.',
  72, 478, 710, 111, { size: 24, color: soft });
phone(s, 'notifications', 'User1 notifications', 895, 142);

s = make('Challenges and issues encountered', '0:55. The original Render database stopped responding. I initialized a replacement and verified access from the Mac with the external TLS endpoint. The simulator was sluggish with two devices running and heavy swap on my MacBook Neo, so I used one simulator. Notifications originally overlapped the notch and then shifted the map, so I moved them behind a bell and corrected the navigation labels. One test setup issue remains. Jest stops before executing tests because it cannot load the current React Navigation modules. The TypeScript check passes.', true);
sections(s, [
  ['Database connectivity', 'Initialized a replacement Render database and verified its external TLS connection.'],
  ['Simulator performance', 'Reduced simulator load by keeping one device running on the MacBook Neo.'],
  ['Notification layout', 'Replaced the popup banner with a bell and a separate notifications screen.'],
  ['Automated test setup', 'TypeScript checks pass. Jest still needs configuration for React Navigation.'],
], 72, 148, 1134, 118, true);

s = make('Remaining work and schedule', '0:45. The schedule is a proposal. In the next phase, I want to run on a physical iPhone and compare GPS and estimated steps on a real route. Then I’ll fix the Jest setup and add focused tests around authentication, session ownership and notifications. Finally I’ll check Android, move credentials to environment variables, prepare database backup and recovery, and rehearse the demonstration. The Week 8 simulator build is ready to show the main flow.', true);
sections(s, [
  ['Week 9  ·  Field testing', 'Test on a physical iPhone and check GPS accuracy and boundary cases.'],
  ['Week 10  ·  Automated checks', 'Fix the Jest setup and test authentication, sessions and notifications.'],
  ['Week 11  ·  Submission preparation', 'Check Android, configure credentials and backups, then rehearse the demo.'],
], 72, 156, 1090, 134, true);
text(s, 'Proposed schedule', 72, 604, 900, 30, { size: 19, bold: true, color: teal });

for (const [i, slide] of deck.slides.items.entries()) {
  const preview = await deck.export({ slide, format: 'png', scale: 0.75 });
  await fs.writeFile(path.join(build, 'slide-' + String(i + 1).padStart(2, '0') + '.png'),
    new Uint8Array(await preview.arrayBuffer()));
}

const draft = path.join(build, 'candidate.pptx');
await (await PresentationFile.exportPptx(deck)).save(draft);
const { finalizePresentation } = await import(pathToFileURL(path.join(toolDir, 'artifact_tool_utils.mjs')).href);
await finalizePresentation({
  explicitTotalSlideCount: 10, workspaceDir: root, candidatePath: draft, finalPath: out,
  pythonExecutable, requiredNativeTableOwnerSlides: [4],
  integrityValidatorPath: path.join(toolDir, 'inspect_presentation_package_integrity.py'),
  layoutValidatorPath: path.join(toolDir, 'inspect_presentation_layout_geometry.py'),
  layoutArgs: ['--expected-slide-size-emu', '12192000,6858000', '--validate-heading-fit', '--require-native-table-slide', '4'],
  fontPolicy: { basis: 'design', families: [font] }, verifyArtifactToolImport: true,
  receiptPath: path.join(build, 'validation.json'),
});
console.log(out);
