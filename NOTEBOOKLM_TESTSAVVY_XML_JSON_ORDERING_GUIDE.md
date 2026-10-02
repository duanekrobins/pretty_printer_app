# NotebookLM Guide: TestSavvy XML and JSON Markdown Payloads

## Purpose

This guide teaches NotebookLM how to read generated TestSavvy Markdown files. Upload this guide together with the generated `.md` files from XML execution payloads, Test Case JSON exports, and Interface Repository JSON exports.

## Supported Payload Types

| Payload Type | Source Shape | Generated Markdown Purpose |
| --- | --- | --- |
| TestSavvy XML Execution Payload | XML containing `automation_sequence`, `scenario`, and `step` nodes | Explains what TestSavvy actually executed in a specific run. This is the closest source to runtime truth. |
| TestSavvy Test Case JSON Export | JSON object containing `automationSequence` and `sequenceList` | Explains the authored test case definition, including active scenarios, excluded scenarios, reusable steps, datasets, custom functions, and selectors. |
| TestSavvy Interface Repository JSON Export | JSON object containing `export.interfaceMapName` and recursive `children` | Explains the interface map tree, logical element paths, actions, selector properties, and locator metadata. |
| TestSavvy Group / Array JSON Export | JSON top-level array | Explains grouped/exported array content and provides a JSON inventory. |

## Most Important Rule: JSON Array Order Is Not Execution Order

For Test Case JSON exports, do not assume `$.sequenceList[]` physical array order is the order TestSavvy executes. The generator sorts scenarios by `sequence`, falling back to `original_sequence`, and preserves the raw JSON array location separately.

Use these sections this way:

| Section | Use For | Do Not Use For |
| --- | --- | --- |
| `## Active Runtime-Equivalent Scenario Index` | Understanding what the test case is expected to execute, sorted by TestSavvy sequence. | Debugging raw export order. |
| `## Active Runtime-Equivalent Scenarios and Steps` | Step-by-step explanation of the active authored flow. | Counting disabled/excluded coverage. |
| `## Excluded / Inactive Scenario Index` | Finding scenarios that exist in the JSON but are excluded from execution. | Answering what the test actually does. |
| `## Excluded / Inactive Scenarios and Steps` | Troubleshooting disabled design content. | Runtime coverage claims. |
| `## Full Raw JSON Scenario Inventory` | Debugging the raw export array and locating JSON paths. | Execution order. |
| `## JSON Inventory` | Finding deeply nested source values. | Human-readable flow analysis unless no other section has the value. |

## Active vs Excluded Execution Rules

A TestSavvy JSON scenario should be considered active/executable only when:

```text
scenario.status is not false
AND scenario.include_status is not false
```

A TestSavvy JSON step should be considered active/executable only when:

```text
parent scenario is active
AND step.status is not false
AND step.include_status is not false
AND stepAssociation.status is not false
```

If any of those values are explicitly false, the generated Markdown marks the row as `Excluded / Not Executed` and records an `Exclusion Reason`, such as `scenario.include_status=false`, `step.include_status=false`, or `stepAssociation.status=false`.

## How To Answer Coverage Questions

When the user asks a coverage question such as:

```text
Which TestSavvy automated test cases click the Submit button?
```

NotebookLM should use `## Submit Button Execution Index` and count only rows where `Execution Status` is `Active / Executable`.

If the user asks:

```text
Which test cases contain Submit, including excluded steps?
```

NotebookLM may include both `Active / Executable` and `Excluded / Not Executed` rows, but it must label them separately.

## Test Case JSON Mapping

| Business Term | JSON Location | Generated Markdown Location | Meaning |
| --- | --- | --- | --- |
| Test Case ID | `$.automationSequence.id` | `## Document Identity`; `## Test Case Summary` | The TestSavvy test case / automation sequence identifier. |
| Test Case Name | `$.automationSequence.name` | Document title; `## Document Identity`; `## Test Case Summary` | Human-readable name of the test case. |
| Raw Scenario List | `$.sequenceList[]` | `## Full Raw JSON Scenario Inventory` | Physical JSON export array. This is not necessarily execution order. |
| Scenario Execution Order | `$.sequenceList[].sequence` | `## Active Runtime-Equivalent Scenario Index` | Primary scenario ordering field. |
| Fallback Scenario Order | `$.sequenceList[].original_sequence` | `## Active Runtime-Equivalent Scenario Index` | Fallback ordering field when `sequence` is missing. |
| Scenario Include Status | `$.sequenceList[].include_status` | Scenario index and scenario metadata tables | False means the scenario exists in JSON but should not be counted as executed coverage. |
| Scenario Status | `$.sequenceList[].status` | Scenario index and scenario metadata tables | False means the scenario is inactive. |
| Scenario Steps | `$.sequenceList[].scenarioFlowList[]` | Scenario step detail tables | Step/action definitions inside each scenario, sorted by step sequence. |
| Step Include Status | `$.sequenceList[].scenarioFlowList[].include_status` | Scenario step detail table `Execution Status` and `Exclusion Reason` | False means the step exists in JSON but is excluded. |
| Step Data Value | `$.sequenceList[].scenarioFlowList[].stepAssociation.value` | Scenario step detail table column `Value` | Data entered, selected, clicked, or passed into a custom function. |
| Step Dataset | `$.sequenceList[].scenarioFlowList[].stepAssociation.dataset_name` | Scenario step detail table column `Dataset` | Dataset used by the step. |
| Step Dataset Header | `$.sequenceList[].scenarioFlowList[].dataset_header_name` | Scenario step detail table column `Dataset Header` | Dataset field/header used by the step. |
| Action Code | `$.sequenceList[].scenarioFlowList[].action_code` | Scenario step detail table column `Action Code`; `## Selector / Action Code Index` | Selenium/WebDriver/custom code for the step. |
| Custom Function | Step `Value` or `Action Code` containing `@...` | `## Custom Functions and Dynamic Values` | TestSavvy custom functions and dynamic values. |

## Test Case JSON Sections

| Section | Meaning |
| --- | --- |
| `## Critical Ordering Rule` | Explicit reminder that raw JSON array order is not execution order. |
| `## Active Runtime-Equivalent Scenario Index` | Active scenario flow sorted by `sequence` / `original_sequence`. |
| `## Excluded / Inactive Scenario Index` | Disabled scenarios, still sorted by sequence for context. |
| `## Full Raw JSON Scenario Inventory` | Raw export array order and JSON path references. |
| `## Submit Button Execution Index` | Purpose-built coverage table for Submit button questions. |
| `## Custom Functions and Dynamic Values` | Dynamic values, `@storeResult`, `@cgifx`, `@advantagefx`, and other custom function usage. |
| `## Selector / Action Code Index` | Locator/action-code analysis table. |
| `## Active Runtime-Equivalent Scenarios and Steps` | Detailed active scenario and step flow. |
| `## Excluded / Inactive Scenarios and Steps` | Detailed disabled scenario and step content. |
| `## JSON Inventory` | Flattened fallback source of truth. |

## XML Execution Payload Mapping

XML execution Markdown files are already runtime ordered because they come from the execution payload.

| Business Term | XML Location | Generated Markdown Location | Meaning |
| --- | --- | --- | --- |
| Test Case ID | `automation_sequence/@id` | `## Execution Summary`; `## Automation Sequence Attributes` | TestSavvy test case / automation sequence identifier. |
| Test Case Name | `automation_sequence/@name` | Title; `## Execution Summary`; `## Automation Sequence Attributes` | Human-readable automated test name. |
| Run ID | `automation_sequence/@run_id` | `## Execution Summary`; `## Automation Sequence Attributes` | Specific execution run. |
| Runtime Scenario Order | `scenario/@count_id` and payload order | `## Scenarios and Steps` | Actual execution order in that run. |
| Runtime Step Order | `step/@count_id` and payload order | Scenario step detail table | Actual step order in that run. |
| Action Code | `step/@action_code` | Scenario step detail table column `Action Code` | Runtime automation code/selectors. |

## Interface Repository JSON Mapping

| Business Term | JSON Location | Generated Markdown Location | Meaning |
| --- | --- | --- | --- |
| Interface Map Name | `$.export.interfaceMapName` | `## Document Identity`; `## Interface Root Summary` | Top-level interface map, such as `FIN`. |
| Interface Element Name | Recursive `interfaceElementName` under `$.export.children[]` | `## Interface Element Index` | Human-readable UI element name. |
| Text Path | Recursive `textPath` under `$.export.children[]` | `## Interface Element Index` | Full logical path to the interface element. |
| ID Path | Recursive `idPath` under `$.export.children[]` | `## Interface Element Index` | TestSavvy internal path/id chain. |
| Actions | Recursive `actions[]` under each element | `## Interface Actions and Properties` | Available action definitions such as Click, Enter, Verify. |
| Selector Properties | `actions[].properties[]` | `## Selector Property Index` | Locator properties such as XPath, ID, name, aria-label, `data-qa`, and `data-qa-id`. |

## Selector Analysis Rules

For Test Case JSON, inspect `Action Code`, `Raw Step JSON`, and `## Selector / Action Code Index`. For Interface Repository JSON, inspect `## Selector Property Index` and `## Interface Actions and Properties`. These sections are the best sources for XPath, `data-qa`, `data-qa-id`, `aria-label`, ID, name, and other locator decisions.

## NotebookLM-Friendly File Names

Generated Markdown filenames are normalized so they contain underscores and end in `.md`.

| Source Type | Filename Pattern |
| --- | --- |
| XML execution payload | `<TestCaseID>_<RunID>.md` |
| Test Case JSON export | `TC_<TestCaseID>.md` |
| Interface Repository JSON export | `IR_<InterfaceMapName>.md` |
| Group JSON export | `GROUP_<source>.md` |

Avoid stale filenames like `998_19586.dat.md`; regenerate them as `998_19586.md`.
