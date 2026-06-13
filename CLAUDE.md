# CLAUDE.md - AI Coding Instructions

## Part 1: Core Coding Rules
1. **Think Before Coding**: State your assumptions clearly. Discuss trade-offs and ask clarifying questions instead of guessing.
2. **Simplicity First**: Write the minimum code required to solve the immediate problem. Avoid speculative features or premature abstractions.
3. **Surgical Changes**: Only modify code directly relevant to the task. Do not "clean up" adjacent code, styling, or comments.
4. **Goal-Driven Execution**: Define what success looks like (e.g., write a test and make it pass). Let the AI iterate until verification.

## Part 2: Agent Orchestration Rules
5. **Keep Deterministic Work out of AI**: Do not make Claude handle raw string formatting or mechanical tasks; delegate these to standard code tools.
6. **Manage Token Budgets**: Enforce strict limits on context usage (e.g., 4k per message, 30k per session) to prevent token bloat.
7. **Resolve Style Conflicts**: If formatting or lint rules conflict, prioritize a single unified configuration and discard the rest.
8. **Verify Context Before Editing**: Always read the surrounding code and imports before writing a single line to ensure compatibility.
9. **Use Business-Logic Tests**: Write meaningful tests that validate actual intent and business outcomes, not just empty code coverage.
10. **Create Step-by-Step Checkpoints**: For multi-step long tasks, halt at milestones to log what was done, what was verified, and what remains.
11. **Match Existing Codebase Style**: Strictly follow established code conventions (e.g., snake_case or class components) even if you disagree.
12. **Explicitly Fail Loud (Fail Loud)**: If a step fails, skips data, or cannot be fully verified, report the error immediately. Never hide uncertainties.
