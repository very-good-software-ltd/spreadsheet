# New React Router App

## Getting Started

### Installation

Install the dependencies:

```bash
npm install
```


### Development

Start the development server:

```bash
npm run dev
```

Your application will be available at `http://localhost:8080`.


## Scripts

Run any of these with `npm run <name>`.

| Script | What it does |
| --- | --- |
| `dev` | Start the development server at `http://localhost:8080`. |
| `build` | Create a production build. |
| `start` | Serve the production build. Run `build` first. |
| `test` | Run the test suite once. |
| `typecheck` | Type-check the project. |
| `lint` | Check formatting and lint rules. |
| `lint:fix` | Apply fixes in place. |
| `lint:routes` | Check the route config with the React Router route linter. |
| `lint:ci` | Run Biome in CI mode. |
| `lint:verify` | Run `lint` and `lint:routes` together. |
| `verify` | Run `lint:verify`, `typecheck` and `test`. Use this before pushing. |


## Building for Production

Create a production build:

```bash
npm run build
```


## Deployment

### Docker Deployment

To build and run using Docker:

```bash
docker build -t my-app .

# Run the container
docker run -p 8080:8080 my-app
```
