# Deal Quill

Bare-bones **Microsoft Word** Office Add-in (task pane) for commercial real estate deal documents.  
Built for **Chad Nasir / RealEstateCA**.

**Demo — not legal advice.** Prompts drive AI behavior; there are no hardcoded contract playbooks or clause libraries.

## Live task pane URL

`https://raw.githack.com/Chadnasir/deal-quill-word/main/taskpane.html`

Manifest SourceLocation points at that URL (local override: `https://localhost:3000/taskpane.html`).

## What it does

1. **Task pane** beside the Word document (Office Add-in).
2. **Strike out** — apply strikethrough to selection (or whole body if none).
3. **Redline** — strike selection and insert a replacement after it.
4. **Give suggestions / Ask AI** — send your prompt + selection/doc context + labeled parties to a **local Grok bridge**; Apply insert or Apply redline.
5. **Parties panel** — on open (and on **Refresh parties**), heuristic detection of names/entities from document text; you label role (buyer, seller, landlord, tenant, lender, borrower, other) and display name.

## First-run UX

The task pane includes a three-step checklist (set the bridge token, select a clause, then strike or ask AI). On first run (no bridge token in Office storage / localStorage), the Grok bridge panel auto-opens and step 1 shows an inline token field so you do not hunt inside collapsed details. After the token is saved, the panel collapses and de-emphasizes again. The checklist disappears permanently after those steps are completed. Empty states guide the next action when there is no document text, the local bridge is offline, or no parties are detected.

Clarity is intentionally not recorded inside Word: Office task panes often block analytics and document text is sensitive. **Clarity primarily on OM portal; Deal Quill uses empty-state onboarding only.**
