<script lang="ts">
    import type { Content, Editor } from '@tiptap/core'
    import {
        EdraEditor,
        EdraToolBar,
        EdraBubbleMenu,
        EdraDragHandleExtended,
    } from '$lib/components/edra/shadcn'

    let content = $state<Content>()
    let editor = $state<Editor>()
    const excludedCommands = ['media']
    function onUpdate() {
        content = editor?.getJSON()
    }
</script>

<div
    class="bg-background z-50 mt-12 size-full max-w-5xl rounded-md border border-dashed"
>
    {#if editor && !editor.isDestroyed}
        <EdraToolBar
            class="bg-secondary/50 flex w-full items-center overflow-x-auto border-b border-dashed p-0.5"
            {editor}
            {excludedCommands}
        />
        <EdraDragHandleExtended {editor} />
        <EdraBubbleMenu {editor} {excludedCommands} />
    {/if}
    <EdraEditor
        bind:editor
        {content}
        class="h-[30rem] max-h-screen overflow-y-scroll pr-2 pl-6"
        {onUpdate}
    />
</div>
