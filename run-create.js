import { spawn } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const child = spawn('node', [
  'agentdemo.js', 'create',
  '--studio-url', 'https://copilotstudio.preview.microsoft.com/environments/f40f796d-2796-ec31-9125-3d11d482dbce/bots/9af8bed6-8bf5-f011-8406-002248e43f6d/overview',
  '--m365-url', 'https://m365.cloud.microsoft/chat/agent/T_9a071d6e-1cee-ac21-6c77-8aebe7ea9e0a.dcc03f21-a85a-40cb-a80d-3659680efa58?auth=2.',
], {
  cwd: __dirname,
  stdio: ['pipe', 'pipe', 'inherit'],
  env: { ...process.env },
});

const answers = [
  { match: 'manually', answer: 'y' },
  { match: 'Agent name', answer: 'Humanitix Event Attendance Agent' },
  { match: 'One-line description', answer: 'An AI agent that helps event organisers track attendance, look up attendees, and manage check-ins using Humanitix event data.' },
  { match: 'instructions', answer: 'This agent connects to Humanitix to provide real-time event attendance tracking. It can search for attendees by name or ticket number, show attendance summaries for specific events, check-in attendees on arrival, provide event statistics including total registrations and check-in rates, and help organisers manage their event logistics through natural language queries in M365 Copilot.' },
  // blank line to end instructions
  { match: '>', answer: '', isInstructions: true },
  // blank line to end platforms
  { match: 'Options:', answer: '' },
  { match: '>', answer: '', isPlatforms: true },
];

let answerIdx = 0;
let buffer = '';
let instructionsDone = false;
let platformsDone = false;

child.stdout.on('data', (data) => {
  const text = data.toString();
  process.stdout.write(text);
  buffer += text;

  if (answerIdx >= answers.length) return;

  const entry = answers[answerIdx];

  if (buffer.includes(entry.match)) {
    setTimeout(() => {
      child.stdin.write(entry.answer + '\n');
      console.log(`\n[AUTO-ANSWER ${answerIdx}]: "${entry.answer}"`);
      answerIdx++;
      buffer = '';

      // After instructions text, send blank line to end multi-line
      if (entry.match === 'instructions') {
        setTimeout(() => {
          child.stdin.write('\n');
          console.log('[AUTO-ANSWER]: (blank line to end instructions)');
          answerIdx++; // skip the next > match
          // Now send blank line for platforms
          setTimeout(() => {
            child.stdin.write('\n');
            console.log('[AUTO-ANSWER]: (blank line to skip platforms)');
            answerIdx = answers.length; // done with manual input
          }, 2000);
        }, 1000);
      }
    }, 500);
  }
});

child.on('close', (code) => {
  console.log(`\n[run-create] Process exited with code ${code}`);
  process.exit(code);
});

child.on('error', (err) => {
  console.error('Failed to start process:', err);
  process.exit(1);
});
