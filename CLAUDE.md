# Toast Performance System — Web App

## Stack & architecture
- Single-file React 18.3.1 app, no build step, loaded via CDN, all JSX written as React.createElement (no JSX transform).
- Tailwind via CDN. Dark theme tokens: bg #09090b, panel #1f1f1f, inset #141414, accent #A66CFF (purple).
- State via useState/useRef/useEffect/useCallback destructured from React.
- Deploys: push to main -> Netlify auto-deploys (toastperformancesystem site). Never go back to manual upload.
- Before committing: extract the <script>...</script> body and run `node --check` on it to catch syntax errors. Nested React.createElement calls are easy to miscount parens on -- prefer breaking large render blocks into small named intermediate variables rather than one giant nested expression.

## Backend
- Google Apps Script web app at cfg.url with cfg.token.
- GET via `tab` query param: no tab/date = strength log, tab=recap, tab=nutrition&date=X (single day), tab=nutrition alone (full history, used for suggestions), tab=weight.
- POST with JSON body: {token, session} | {token, meal} | {token, weight} | {token, deleteMeal}.
- Known gap: no access to the Apps Script source from the cloud assistant side. Can't directly confirm it persists/returns a meal_type column, or that tab=nutrition (no date) genuinely returns full history.

## Nutrition logging (current design -- do not revert to the old timeline)
- 4 meal-type buckets: Breakfast / Lunch / Dinner / Snack (MEAL_TYPES, DEFAULT_MEAL_TYPE = "Lunch"). New/typed/estimated entries default to Lunch.
- Drag the grip handle on a logged item to move it between buckets, via native Pointer Events (not HTML5 DnD).
- mealType persists via pushMeal -> POST; reads back via mapMealRow (r.meal_type).
- Previous design was an 8am-10pm hourly timeline -- explicitly replaced with the 4-bucket version. Don't bring it back.

## OPEN ISSUE -- Suggested meals / Quick Add look stale
mealCombos groups history by date+mealType, finds 2+-food combos repeated on 2+ occasions, ranks by same-weekday then total frequency, shows top 6. Requires fetchAllMeals (tab=nutrition, no date) to actually return full history. User reports Quick Add/Suggested Meals look like the same static list, not reflecting recent/real logs. Not yet root-caused -- check: stale cached deploy vs. fetchAllMeals actually returning full history vs. not enough repeated combos logged yet.

## OPEN ISSUE -- Floating numeric-keypad toggle unconfirmed on device
Weight/reps inputs should default to a numeric keypad (inputMode=decimal/numeric) with a way to switch to a full keyboard for things like "BW" or "AMRAP". Current implementation is a floating toolbar pinned above the iOS keyboard using window.visualViewport offset math, with a 123/ABC toggle that blurs/refocuses the input to force inputMode to re-apply. Not yet confirmed working on a real device/live deploy -- verify this first.

## Workflow notes
- Preview before live: never push app changes straight to main. Push to a branch, open a PR, wait for the Netlify deploy preview (deploy-preview-<N>--toastperformancesystem.netlify.app), share that URL with the user, and merge to main only after they confirm it looks right. Docs-only changes (e.g. this file) can go straight to main.
- Preview caveats: previews are a different origin, so Settings (Sheet URL/token) saved on the live site don't carry over -- re-enter them on the preview to see real data. Non-production deploys require Netlify team login.
- Pushes happen exclusively via local Claude Code -> GitHub -> Netlify auto-deploy.
- GitHub auth: fine-grained PAT scoped to this repo only, Contents: Read/write.
- Keep styling consistent with the dark theme tokens and "bebas" header class across the app.
