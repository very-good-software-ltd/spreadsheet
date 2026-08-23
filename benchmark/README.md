# Benchmarks

The scripts in this directory measure how long this library takes to read and write a spreadsheet, and how much memory it uses doing it.
They run `exceljs` and `xlsx` over the same files so you can compare.


## Running the benchmarks

```shell
npm run benchmark        # run read and write benchmarks
npm run benchmark:read   # run read benchmarks
npm run benchmark:write  # run write benchmarks
```

Each one builds the library first, so you are always measuring the code in your working tree.

Four options change what a run does.

- `--write=100000` writes that many rows instead of the default million.
  Pass a list like `--write=100000,1000000` to see whether memory moves with the row count.
- `--cap=150` runs each measurement under a heap limit of that many megabytes.
  See "Reading the memory column" below for why you would want this.
- `--timeout=600` gives each measurement that many seconds before it is killed and reported as timed out.
  The default is 120.
  Pass `--timeout=0` to wait however long it takes.
- Any argument that is not an option is treated as a file path, so you can measure one file instead of the whole directory.

Options go after a `--` when you use the npm scripts, e.g. `npm run benchmark:read -- --cap=512`.


## How a measurement is taken

Every measurement runs in its own Node process, one per library, mode and file.
That is the whole reason the scripts are split into a runner and two workers.
If all the libraries ran in one process, the garbage left by the first would still be there when the second started, and the memory figures would say more about the order we ran them in than about the libraries.

`run.mjs` is the runner.
It works out which combinations to measure, starts each one, and prints the table.
`read-file.mjs` and `write-file.mjs` are the workers.
Each one does a single measurement and prints one line of JSON with the time, the amount of work done and the peak memory.

A worker reads or writes every cell rather than only opening the file.
Otherwise a library could look fast by leaving the parsing until someone asks for a value.

The write workers all write to a real file on disk.
Nothing collects the output in memory, so what you measure is the library rather than where the bytes end up.


## What the modes mean

Reading has two modes.
In `stream` mode the worker reads one row at a time and throws it away, which is the path that is meant to hold a fixed amount of memory.
In `load` mode the worker keeps every row, which is the path that holds the whole file.

Writing has the same two, plus two more.
In `stream` mode the worker hands the rows over as a source and lets the writer pull them.
In `load` mode the worker builds every row in memory first.
`region-stream` and `region-load` fill a named region in a template, which moves the rows below it, and they are ours alone because no other library does that.

Not every library appears in every row of the table.
`xlsx` has no streaming read, so it only ever runs in `load` mode.
`exceljs` reads `.xlsx` and not `.ods`, so it is left out for those files.


## Reading the memory column

The memory column is peak resident memory, taken from the operating system.
It counts every page the process ever held, including memory the runtime has finished with but not yet handed back.
So a figure can be higher than the amount the library actually needed at any one moment.

`--cap` is the answer to that.
Under a heap limit, a library that genuinely needs more than the limit runs out of memory and says so, and one that does not finishes.
That tells you what is actually held.
The catch is that a limit applies to every library in the run, so a limit low enough to be interesting for one file will fail every larger one.
Use it to answer a specific question, not as a default.


## Files

The benchmarks read whatever is in the `files` directory.
There is no fixed set, so you can drop any `.xlsx` or `.ods` file in there and it is picked up on the next run.

Prefix a name with an underscore to leave it out, e.g. `_the-file.xlsx` is skipped and `the-file.xlsx` is used.
That way you can keep a collection locally and still run against one file at a time.

The files are not committed.
`files/.gitignore` ignores everything in the directory, and it is there on purpose.
The ones we use are tens of megabytes each, and the largest is over 40 MB.
Committing them would make every clone carry all of them forever, which is a high price for data anyone can download.

The files we have used come from the police recorded crime open data published by the UK government.

https://www.gov.uk/government/statistical-data-sets/police-recorded-crime-and-outcomes-open-data-tables

The outcomes open data tables are the useful ones, because a single sheet in them runs to millions of cells.
`prc-outcomes-open-data-mar2021-tables-240425.xlsx` is the largest we have tried, at 40 MB on disk and 8.7 million cells.
Any spreadsheet works, so if you have a file of your own that is slow, that is the one worth measuring.
