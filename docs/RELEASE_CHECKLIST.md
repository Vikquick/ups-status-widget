# Release Checklist

## Pre-release

- [ ] `main` is green in GitHub Actions.
- [ ] Open high-priority issues for target milestone are resolved.
- [ ] Version/release notes are reviewed and factual.

## Build Verification

- [ ] Build workflow produced x64 and x86 artifacts.
- [ ] App launches on Windows and shows live status.
- [ ] Manual smoke test: tray actions, startup toggle, status endpoints.

## Tag and Publish

- [ ] Create annotated tag (example): `git tag -a v2.0.0 -m "Release v2.0.0"`.
- [ ] Push tag: `git push origin v2.0.0`.
- [ ] Confirm GitHub Release has both binaries attached.

## Post-release

- [ ] Verify release page and README links.
- [ ] Confirm update-check endpoint/version metadata is reachable.
- [ ] Move milestone to closed and triage follow-up issues.
