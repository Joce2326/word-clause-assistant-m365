/// <reference types="office-js" />

export async function insertTextIntoWord(text: string): Promise<void> {
  await Word.run(async (context) => {
    context.document.body.insertParagraph(text, Word.InsertLocation.end);
    await context.sync();
  });
}