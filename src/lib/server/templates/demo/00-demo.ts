import { Document } from 'docx'
import { DocxBuilder } from '../builder'

async function arbitraryAwaitable(output: string): Promise<string> {
    return output
}

export async function demoDoc(): Promise<Document> {
    const builder = new DocxBuilder()
    const doc = await builder.document({
        sections: [
            builder.section({
                children: [
                    builder.Paragraph({ text: 'Hello world!' }),
                    builder.Paragraph(
                        arbitraryAwaitable('My awaited string output.')
                    ),
                    builder.Paragraph({
                        text: 'This will not appear.',
                        predicate: () => {
                            return false
                        },
                    }),

                    builder.Paragraph({
                        children: [
                            builder.llm.textRun({
                                prompt: 'this is the test prompt',
                            }),
                        ],
                    }),
                    builder.Paragraph({
                        children: [
                            builder.TextRun({
                                text: 'No.',
                            }),
                        ],
                    }),
                ],
            }),
        ],
    })
    return doc
}
