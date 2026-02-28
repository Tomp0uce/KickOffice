## Description

This PR introduces critical advanced capabilities, fundamentally bridging the feature gap between KickOffice and OpenExcel while expanding these features across the entire Office suite.

### Key Features:

1. **Secure Dynamic Execution (`ses` Sandbox)**:
   - Introduced `sandbox.ts` and `lockdown.ts`.
   - Added new escape-hatch tools (`eval_officejs`, `eval_wordjs`, `eval_powerpointjs`, `eval_outlookjs`) permitting the agent to run dynamically generated JavaScript safely inside a compartment.
2. **File Processing via Backend**:
   - Added the `/api/upload` route integrating `multer`, `pdf-parse`, `mammoth`, and `xlsx` for parsing local files entirely securely on the backend.
   - Files are automatically sent to the LLM context through `<attachments>`, enabling it to synthesize external datasets into Office files.
3. **Advanced Excel Tools (Ported from OpenExcel)**:
   - `findData`: Regex and case-sensitive cross-sheet search.
   - `duplicateWorksheet`, `hideUnhideRowColumn`.
   - `getAllObjects`: Chart and pivot discovery.
   - `modifyObject`: Chart and pivot deletion operations.
4. **Prompt Optimization**:
   - Upgraded system prompts enforcing `batchProcessRange` to drastically reduce iterative overhead from `set_cell_range`, optimizing large dataset transforms.
   - Documented explicit overwrite warnings to safeguard user data.

### Documentation Updates

- `README.md`, `agents.md`, and `CHANGELOG.md` have been fully updated to reflect the new feature implementations, configuration requirements, and API additions.

### Impacts Validated

- New tool additions have been implicitly accommodated by `toolStorage.ts` automatically marking them as enabled by default.
- UI features seamlessly ignore UI-breaking loops. All changes pass the TS compiler check (`tsc --noEmit`).
