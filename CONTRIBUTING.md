# Contributing to Office File Viewer

Thank you for helping improve `office-file-viewer`. Contributions are welcome in the form of bug reports, documentation improvements, compatibility reports, and pull requests.

## Before opening an issue

- Check the existing issues and the documented [limitations](https://gyxing.github.io/office-file-viewer/docs#limitations).
- For document rendering problems, include the file format, browser, React version, operating system, and a minimal reproducible example when possible.
- Do not upload confidential documents or sensitive personal data to a public issue.

## Local development

```bash
yarn
yarn start
```

Run the complete validation before opening a pull request:

```bash
yarn run check
```

## Pull requests

- Keep each pull request focused on one problem or feature.
- Update the English and Chinese documentation when public behavior changes.
- Add or update a reproducible sample when a change affects a supported Office format.
- Explain compatibility, performance, security, and rendering trade-offs in the pull request description.
- Ensure the existing type, lint, style, build, and compatibility checks pass.

By participating, you agree to follow the [Code of Conduct](./CODE_OF_CONDUCT.md).
