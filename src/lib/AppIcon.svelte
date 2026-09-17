<script lang="ts">
  import { appInfo, type AppKind } from '../model';
  let { kind, small = false }: { kind: AppKind; small?: boolean } = $props();
</script>

<span class="app-icon {kind}" class:small aria-hidden="true">{appInfo[kind].letter}<i></i></span>
