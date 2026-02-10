import { createClient } from '@papert-code/sdk-typescript';

async function main() {
  const client = createClient({
    model: 'gpt-5.2', permissionMode: 'yolo', skillsPath: ['.papert/skills'], debug: true, 
    stderr: (line) => console.error("[papert]", line),
  });

  const session = client.createSession({ sessionId: 'sdk-pdf_to_ppt' })

  const query = 'Use the pptx skill for this task, Task: make the south india banks quartly report south_india_bank_quarter_3_2025-2026_financial_result.pdf and create the beautiful PPT to submit to south india bank stake holders use the PPtX skill`, do not ask any question, choice what is best u think'

  const pdf_to_ppt = await session.send(query);
  
  for (const msg of [...pdf_to_ppt]) {
    if (msg.type === 'assistant') {
      console.log('\nASSISTANT:\n', msg.message.content);
    }
    if (msg.type === 'result') {
      console.log('\nRESULT:', msg.subtype, msg.is_error ? 'ERROR' : 'OK');
    }
  }

  await client.close();
}

main().catch((err) => { console.error(err); process.exit(1); });
