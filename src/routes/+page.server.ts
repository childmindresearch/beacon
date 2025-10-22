import { demoDoc } from '$lib/server/templates/demo/00-demo'
import { Packer } from 'docx'
import type { PageServerLoad } from './$types'

export const load: PageServerLoad = async () => {
    const doc = await demoDoc()
    const base64 = await Packer.toBase64String(doc)
    return { doc: base64 }
}
