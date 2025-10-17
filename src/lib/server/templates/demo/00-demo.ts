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
                    builder.paragraph({ text: 'Hello world!' }),
                    builder.paragraph(
                        arbitraryAwaitable('My awaited string output.')
                    ),
                    builder.paragraph({
                        text: 'This will not appear.',
                        predicate: () => {
                            return false
                        },
                    }),

                    builder.paragraph({
                        children: [
                            builder.llm.textRun({
                                prompt: 'this is the test prompt',
                            }),
                        ],
                    }),
                    builder.paragraph({
                        children: [
                            builder.textRun({
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
