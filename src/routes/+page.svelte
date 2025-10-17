<script lang="ts">
    import { onMount } from 'svelte'

    const { data } = $props()

    onMount(async () => {
        const doc = data.doc

        const byteCharacters = atob(doc)
        const byteNumbers = new Array(byteCharacters.length)
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i)
        }
        const byteArray = new Uint8Array(byteNumbers)
        const blob = new Blob([byteArray], {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        })

        const url = URL.createObjectURL(blob)
        const link = document.createElement('a')
        link.href = url
        link.download = 'template.docx'

        document.body.appendChild(link)
        link.click()

        document.body.removeChild(link)
        URL.revokeObjectURL(url)
    })
</script>

hi
