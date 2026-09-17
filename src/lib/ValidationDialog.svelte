<script lang="ts">
  import { untrack } from 'svelte';
  import Modal from './Modal.svelte';
  import { validationTypes, validationOperators, type ValidationRule } from '../sheet-validation';
  let {
    rule,
    count,
    onclose,
    onapply,
  }: {
    rule?: ValidationRule;
    count: number;
    onclose: () => void;
    onapply: (rule: ValidationRule) => void;
  } = $props();
  const initial = untrack(() => rule);
  let type = $state(initial?.type || 'whole'),
    operator = $state(initial?.operator || 'between');
  let first = $state(String(initial?.formulae[0] ?? '1')),
    second = $state(String(initial?.formulae[1] ?? '10'));
  let allowBlank = $state(initial?.allowBlank ?? true),
    showError = $state(initial?.showErrorMessage ?? true);
  let prompt = $state(initial?.prompt || ''),
    errorMessage = $state(initial?.error || ''),
    error = $state('');
  let promptTitle = $state(initial?.promptTitle || ''),
    errorTitle = $state(initial?.errorTitle || '');
  let showInput = $state(initial?.showInputMessage ?? true);
  function apply(clear = false) {
    try {
      onapply(
        clear
          ? { type: 'none', formulae: [] }
          : {
              type,
              operator: type === 'list' ? undefined : operator,
              formulae: [first, second],
              allowBlank,
              showErrorMessage: showError,
              errorStyle: 'stop',
              prompt,
              error: errorMessage,
              promptTitle,
              errorTitle,
              showInputMessage: showInput,
            },
      );
    } catch (e) {
      error = (e as Error).message;
    }
  }
</script>

<Modal title="Data validation" {onclose} wide>
  <p class="muted">
    Apply to {count} selected cell{count === 1 ? '' : 's'}. Relative references start at the first
    cell. Existing values are retained.
  </p>
  <div class="validation-columns">
    <section aria-label="Validation criteria">
      <h3>Criteria</h3>
      <div class="criteria">
        <label class="field"
          ><span>Allow</span><select aria-label="Validation type" bind:value={type}>
            {#if !validationTypes.includes(type as (typeof validationTypes)[number])}<option
                value={type}>{type} (choose a supported rule to edit)</option
              >{/if}
            {#each validationTypes as choice}<option value={choice}
                >{{
                  whole: 'Whole number',
                  decimal: 'Decimal',
                  list: 'List',
                  textLength: 'Text length',
                }[choice]}</option
              >{/each}
          </select></label
        >
        {#if type !== 'list'}<label class="field"
            ><span>Comparison</span><select
              aria-label="Validation comparison"
              bind:value={operator}
            >
              {#each validationOperators as choice}<option value={choice}
                  >{{
                    between: 'Between',
                    notBetween: 'Not between',
                    equal: 'Equal to',
                    notEqual: 'Not equal to',
                    greaterThan: 'Greater than',
                    lessThan: 'Less than',
                    greaterThanOrEqual: 'Greater than or equal to',
                    lessThanOrEqual: 'Less than or equal to',
                  }[choice]}</option
                >{/each}
            </select></label
          >{/if}
      </div>
      <label class="field"
        ><span>{type === 'list' ? 'Source' : 'Value or minimum'}</span><input
          aria-label="Validation first value"
          bind:value={first}
          maxlength="256"
        /></label
      >
      {#if type !== 'list' && ['between', 'notBetween'].includes(operator)}<label class="field"
          ><span>Maximum</span><input
            aria-label="Validation second value"
            bind:value={second}
            maxlength="256"
          /></label
        >{/if}
      {#if type === 'list'}<p class="muted">
          Enter a quoted list such as "Red,Green,Blue", a range such as $A$1:$A$3, or a defined
          name.
        </p>{/if}
      <label><input type="checkbox" bind:checked={allowBlank} /> Ignore blank</label>
    </section>
    <section aria-label="Validation messages">
      <h3>Input message</h3>
      <label><input type="checkbox" bind:checked={showInput} /> Show input message</label>
      <label class="field"
        ><span>Title</span><input
          aria-label="Validation input title"
          bind:value={promptTitle}
          maxlength="32"
        /></label
      >
      <label class="field"
        ><span>Message</span><input
          aria-label="Validation input message"
          bind:value={prompt}
          maxlength="255"
        /></label
      >
      <h3>Error alert</h3>
      <label
        ><input type="checkbox" bind:checked={showError} /> Reject invalid entries with a Stop alert</label
      >
      <label class="field"
        ><span>Title</span><input
          aria-label="Validation error title"
          bind:value={errorTitle}
          maxlength="32"
        /></label
      >
      <label class="field"
        ><span>Message</span><input
          aria-label="Validation error message"
          bind:value={errorMessage}
          maxlength="225"
        /></label
      >
    </section>
  </div>
  {#if error}<p role="alert">{error}</p>{/if}
  <div class="modal-actions">
    <button onclick={() => apply(true)}>Clear validation</button><button onclick={onclose}
      >Cancel</button
    ><button class="primary" onclick={() => apply()}>Apply validation</button>
  </div>
</Modal>

<style>
  .validation-columns {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(min(100%, 280px), 1fr));
    gap: 24px;
  }
  .criteria {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(min(100%, 180px), 1fr));
    gap: 12px;
  }
  h3 {
    font-size: 14px;
    margin: 16px 0 10px;
  }
  label:not(.field) {
    display: flex;
    gap: 8px;
    align-items: center;
    font-size: 13px;
  }
</style>
