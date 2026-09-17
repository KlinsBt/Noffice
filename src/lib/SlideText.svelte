<script lang="ts">
  import type { SlideElement } from '../model';
  import { fontStack } from '../fonts';
  import { slideTextParagraphs } from '../slide-text';
  let { element }: { element: SlideElement } = $props();
  let paragraphs = $derived(slideTextParagraphs(element));
</script>

{#each paragraphs as runs}
  <div class="slide-text-paragraph" style:text-align={runs[0]?.align || element.align}>
    {#each runs as run}<span
        style:font-family={fontStack(run.fontFamily)}
        style:font-size="{run.fontSize}px"
        style:font-weight={run.bold ? 700 : 400}
        style:font-style={run.italic ? 'italic' : 'normal'}
        style:text-decoration={run.underline ? 'underline' : 'none'}
        style:color={run.color}>{run.text}</span
      >{/each}
    {#if !runs.some((r) => r.text)}<br />{/if}
  </div>
{/each}

<style>
  .slide-text-paragraph {
    white-space: pre-wrap;
    overflow-wrap: break-word;
    line-height: 1.2;
  }
</style>
