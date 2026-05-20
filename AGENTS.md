# Agent Rule Book

These rules apply by default for all future development work unless the user explicitly overrides them.

## 1. Build Structure Before Complexity
- Do not keep adding behavior into a file or module that already mixes multiple responsibilities.
- When a change would deepen coupling, extract the boundary first or in the same change.
- Keep entrypoints thin. They should compose flows, not contain deep implementation logic.

## 2. Enforce Clear Boundaries
- Separate domain logic, IO, orchestration, rendering, and configuration into explicit modules.
- Keep shared contracts and typed payloads in stable, reusable boundaries.
- Avoid anonymous data structures when a typed or explicit contract is warranted.
- Do not let presentation formats become the source of truth for business behavior.

## 3. Keep Product Logic Out Of UI Glue
- Do not hide business rules inside rendering code, widget callbacks, formatting helpers, or copy.
- UI code may present, format, filter, and group data, but ranking, scoring, policy, and recommendation logic must live outside it.
- If a UI change requires product logic, define the logic in a testable layer first.

## 4. Refactor Before Expanding
- If a requested feature would materially increase architecture debt, stop and extract the right boundary first.
- Do not use “just one more feature” as a reason to postpone necessary cleanup.
- Prefer structural safety over short-term implementation speed when the tradeoff matters.

## 5. One Change, One Primary Purpose
- Each change should have one dominant purpose: refactor, feature, bug fix, performance fix, test hardening, or documentation.
- Do not mix unrelated concerns in the same change without a concrete reason.
- If refactor and feature work must happen together, do the structural move first and validate it before layering behavior.

## 6. Tests Are Required For Behavior Changes
- Any change to user-visible behavior, ranking, decision logic, generation logic, filtering, or data handling must include automated regression coverage in the same change.
- Prefer deterministic unit tests for pure logic.
- Keep integration or browser smoke coverage green for user-facing flows.
- Do not rely on manual validation alone when behavior changes.

## 7. Validate Runtime Cost Early
- Review network, disk, and compute cost before merging behavior changes.
- Avoid repeated identical fetches, duplicated work, hidden retries, and unnecessary recomputation.
- Prefer one fetch plus local reuse over repeated calls for the same payload.
- Test harnesses must run against isolated infrastructure and must not silently reuse unrelated local sessions.

## 8. Preserve Momentum With Small, Verifiable Steps
- Break complex work into steps that can be verified independently.
- Prefer reversible, validated increments over wide speculative rewrites.
- Do not leave the codebase in a half-moved state longer than necessary around a policy boundary.

## 9. Documentation Is Part Of Delivery
- Update roadmap, truth, handoff, and execution documents whenever architecture, behavior, validation status, or sequencing changes materially.
- Record what changed, what remains deferred, why it is deferred, and what the next safe step is.
- If the implementation reality and the documentation disagree, fix the documentation in the same change.

## 10. Define Done Rigorously
- Code is in the correct boundary.
- The changed behavior is covered by tests.
- Existing validation passes.
- Documentation reflects the new state.
- The change does not make the next change harder.

## 11. Use Explicit Stop Conditions
- If the boundary is unclear, stop and define it.
- If the data contract is unclear, stop and define it.
- If the test plan is unclear, stop and define it.
- If a change cannot be explained in a clean architecture narrative, it is not ready to merge.

## 12. Default Bias
- Prefer fewer, cleaner changes over faster layered changes.
- Prefer explicit contracts over implicit coupling.
- Prefer extraction before expansion.
- Prefer maintainability over cleverness.
- Prefer correctness and traceability over speed when those goals conflict.

## 13. Product And UX Decision Quality
- Judge product, UX, and design work by whether the user can complete the core task quickly and correctly, not by whether a component renders cleanly or looks improved.
- If the user still cannot infer the next action and reason for it in a few seconds, treat the feature as failing and redesign the concept instead of polishing the presentation.
- Do not present alternatives unless they are serious contenders; if one path is clearly superior, recommend only that path.

## 14. Project-Level Extension Rule
- A repository may add project-specific rules below this rule book.
- Project-specific rules may specialize these defaults but should not weaken them unless explicitly intended and documented.
