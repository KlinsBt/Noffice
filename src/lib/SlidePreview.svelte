<script lang="ts">
  import { fontStack } from '../fonts';
  import { paintCSS } from '../drawing-colors';
  import type { Slide } from '../model';
  import SlideOutline from './SlideOutline.svelte';
  import SlideText from './SlideText.svelte';
  let { slide, aspectRatio = 16 / 9 }: { slide: Slide; aspectRatio?: number } = $props();
  let height = $derived(960 / aspectRatio);
  let verticalScale = $derived(height / 540);
</script>

<svg viewBox={`0 0 960 ${height}`} class="slide-preview" aria-label="Slide preview">
  <rect width="960" {height} fill={paintCSS(slide.background, slide.backgroundOpacity)} />
  {#each slide.elements as el (el.id)}
    <g
      transform={`rotate(${el.rotation || 0} ${el.x + el.w / 2} ${(el.y + el.h / 2) * verticalScale})`}
    >
      {#if el.type === 'image'}<image
          href={el.src}
          x={el.x}
          y={el.y * verticalScale}
          width={el.w}
          height={el.h * verticalScale}
          preserveAspectRatio="none"
          transform={`translate(${el.flipH ? 2 * el.x + el.w : 0} ${el.flipV ? (2 * el.y + el.h) * verticalScale : 0}) scale(${el.flipH ? -1 : 1} ${el.flipV ? -1 : 1})`}
        />
      {:else if el.type === 'rect'}<rect
          x={el.x}
          y={el.y * verticalScale}
          width={el.w}
          height={el.h * verticalScale}
          fill={paintCSS(el.fill, el.fillOpacity)}
        />
      {:else if el.type === 'ellipse'}<ellipse
          cx={el.x + el.w / 2}
          cy={(el.y + el.h / 2) * verticalScale}
          rx={el.w / 2}
          ry={(el.h / 2) * verticalScale}
          fill={paintCSS(el.fill, el.fillOpacity)}
        />
      {:else}<foreignObject
          x={el.x}
          y={el.y * verticalScale}
          width={el.w}
          height={el.h * verticalScale}
          ><div
            style:background={paintCSS(el.fill, el.fillOpacity)}
            style:height="100%"
            style:color={el.color}
            style:font-family={fontStack(el.fontFamily)}
            style:font-style={el.italic ? 'italic' : 'normal'}
            style:text-decoration={el.sourceText ? 'none' : el.underline ? 'underline' : 'none'}
            style:font-size="{el.fontSize}px"
            style:font-weight={el.bold ? 700 : 400}
            style:text-align={el.align}
            style="white-space:pre-wrap;overflow-wrap:break-word;line-height:1.2"
          >
            <SlideText element={el} />
          </div></foreignObject
        >{/if}
      <SlideOutline
        outline={el.outline}
        ellipse={el.type === 'ellipse'}
        x={el.x}
        y={el.y * verticalScale}
        width={el.w}
        height={el.h * verticalScale}
      />
    </g>
  {/each}
</svg>
