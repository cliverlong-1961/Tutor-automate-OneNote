# Open questions and verification

## Deployment

- [ ] Inspect `.github/workflows/` and record the exact frontend deployment trigger, build command and output directory.
- [ ] Confirm GitHub Pages publishing configuration and custom-domain setup.
- [ ] Confirm how Azure Container Apps receives new backend images and whether deployment is automatic.
- [ ] Record Azure resource group, Container App name, region, registry and image-tag strategy.
- [ ] Confirm how production secrets and environment variables are managed.
- [ ] Establish whether commits to `development` or another branch trigger deployment.

## OneNote and Microsoft Graph

- [ ] Verify all notebook and section loading functions and their error handling.
- [ ] Confirm app registration permissions and consent configuration.
- [ ] Decide whether anonymous sharing links are acceptable for student recordings or whether a restricted sharing scope is needed.
- [ ] Test recovery from partial preparation/page-creation failures.

## File handling

- [ ] Decide whether MKV recordings should appear in the recognised-lessons panel. The inspected grouping code filters for MP4, although the picker lists MP4 and MKV.
- [ ] Confirm transcript naming conventions.
- [ ] Confirm exact PDF filename suffix from the backend.
- [ ] Confirm lesson status handles all supported formats and naming patterns.

## Backend

- [ ] Inspect the complete `backend/main.py` and verify all endpoint request/response models.
- [ ] Inspect PDF rendering implementation, dependencies, fonts and mathematical typesetting.
- [ ] Review obsolete Ollama references in `backend/.env.example`, `backend/README.txt` and `Run-Backend.bat`. The inspected active summary workflow uses Gemini, not local Ollama.
- [ ] Verify deployed Gemini model name and availability.

## Security and operations

- [ ] Check whether `backend/.env` or other secret-bearing files were ever committed to Git history. `.gitignore` alone does not prove they were never committed.
- [ ] Confirm production CORS origins.
- [ ] Document recording retention, deletion and access-control expectations.
- [ ] Check whether logs can contain personal data or lesson content.

Update the relevant documentation and remove or mark checklist items when resolved.
