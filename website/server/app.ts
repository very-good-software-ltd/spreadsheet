import { createRequestHandler } from "@react-router/express";
import express from "express";
import { RouterContextProvider } from "react-router";

import { exampleContext } from "~/context";

export const app = express();

app.use(
  createRequestHandler({
    build: () => import("virtual:react-router/server-build"),
    getLoadContext() {
      const context = new RouterContextProvider();

      context.set(exampleContext, "Example Context Value");

      return context;
    },
  }),
);
