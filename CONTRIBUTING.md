# Contributing

Change src/ or an owning adapter, never generated dist files. Preserve public
exports and compatibility entry points unless a versioned change removes them.
Keep ERP business rules out of the generic library.

Use focused Node tests for values/coordinates, serialized workbook checks for
Excel output, and browser fixtures for DOM geometry and images. Every bug fix
should reproduce the original problem. Fixtures must be synthetic; never commit
customer workbooks, signatures, company stamps, credentials or local reports.

Before a pull request:

1. Run npm test.
2. Run npm run build and npm run build:check.
3. Run npm run test:browser on Chrome, Edge and Firefox.
4. Update affected README sections and CHANGELOG.
5. Review git diff --check and staged files.

The CI badge, GitHub Issues, CHANGELOG and Releases are the maintenance record.
Do not add a separate monitoring service for this browser library. Development
dependencies are lockfile-pinned; consumer applications load ExcelJS separately.

## Preview releases

Tag-triggered CI repeats tests before publication. The release job requires a
successful check job and a version matching package.json and the dist manifest.
Version v0.1.0 is a GitHub prerelease.

~~~sh
npm run build
npm run check
git diff --check
git tag -a v0.1.0 -m "HTML Export to XLSX JS 0.1.0 preview"
git push origin main
git push origin v0.1.0
~~~

Release assets are the readable/minified exporter, checksum manifest and MIT
license. ExcelJS is supplied separately. Consumers should select a fixed version,
keep the previous asset for rollback and verify actual forms before deployment.
Production assets must not be overwritten by library tests.

Report bugs with minimal synthetic HTML, browser/version, export options, expected
cells and actual cells. A screenshot helps diagnose layout but cannot establish
that serialized workbook values are correct.
