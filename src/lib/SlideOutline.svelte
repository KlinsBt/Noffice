<script lang="ts">
  import { paintCSS } from '../drawing-colors';
  import { outlineDashArray, type SlideOutline } from '../slide-outline';
  let {
    outline,
    ellipse = false,
    x = 0,
    y = 0,
    width,
    height,
  }: {
    outline?: SlideOutline;
    ellipse?: boolean;
    x?: number;
    y?: number;
    width: number;
    height: number;
  } = $props();
</script>

{#if outline && outline.width > 0 && outline.color !== 'transparent' && outline.opacity > 0}
  <g
    class="object-outline"
    fill="none"
    stroke={paintCSS(outline.color, outline.opacity)}
    stroke-width={outline.width}
    stroke-dasharray={outlineDashArray(outline)}
    pointer-events="none"
  >
    {#if ellipse}<ellipse cx={x + width / 2} cy={y + height / 2} rx={width / 2} ry={height / 2} />
    {:else}<rect {x} {y} {width} {height} />{/if}
  </g>
{/if}
