# ECPI Augmented Craftsman Faculty Conference

This repository contains an eight-slide presentation for ECPI University's November 2026 Faculty Conference. It introduces the Augmented Craftsman Model and the future of software development education, including program progress, a development roadmap, and opportunities for faculty partnership.

The presentation is available as an interactive, browser-based deck and as a downloadable PowerPoint file.

## View the presentation

Open [`index.html`](./index.html) in a browser, or serve the repository locally:

```bash
python3 -m http.server 8000
```

Then visit [http://localhost:8000](http://localhost:8000).

Navigate with the on-screen arrows or keyboard:

- **Right arrow**, **Page Down**, or **Space**: next slide
- **Left arrow** or **Page Up**: previous slide
- **Home** / **End**: first / last slide
- On touch screens, swipe left or right

Slide URLs use hashes such as `#slide-3`. The footer's **Download PowerPoint** link downloads `ECPI_Augmented_Craftsman_Faculty_Conference_2026.pptx`.

## Build the PowerPoint

The PowerPoint is generated from `build_presentation.js` using Node.js and npm:

```bash
npm ci
npm run build
```

The build writes `ECPI_Augmented_Craftsman_Faculty_Conference_2026.pptx` in the repository root, replacing the existing file.

## Repository contents

- [`index.html`](./index.html) — interactive, responsive browser presentation and slide navigation.
- [`build_presentation.js`](./build_presentation.js) — PowerPoint generation script, including speaker notes.
- [`ECPI_Augmented_Craftsman_Faculty_Conference_2026.pptx`](./ECPI_Augmented_Craftsman_Faculty_Conference_2026.pptx) — generated, downloadable presentation.
- [`Phase_One_Progress_Infographic.png`](./Phase_One_Progress_Infographic.png) and [`Phase_Two_Progress_Infographic.png`](./Phase_Two_Progress_Infographic.png) — phase progress infographics.
- [`package.json`](./package.json) and [`package-lock.json`](./package-lock.json) — Node.js dependency and build script configuration.

## License

This project is distributed under the MIT License. See [`LICENSE`](./LICENSE).
