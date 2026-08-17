# very-good-spreadsheet

## Excel fidelity

When code touches Excel behaviour, two things must hold.
It handles what Excel actually does, edge cases included, not just the happy path.
And every quirk it relies on is documented where it lives.

Do not invent quirks.
If unsure a behaviour is real, confirm it against ECMA-376 or a real file, or leave it as an open question.
A confidently wrong quirk is worse than an open one.

Pin value conversions with tests against known real values, not a round-trip fixture.
A fixture that encodes and decodes with the same constant cannot catch a wrong constant.
`serialToDate` is the example, its tests use known Excel serials.

Document a quirk to the bar set by `src/xlsx/date.ts`: the surprising constant, the spec reason, at the point of use.

The `excel-fidelity` agent audits both duties and shares this bar.


## OOXML vocabulary

Never litter the parse code with raw OOXML identifiers.
We cannot avoid their spec strings, but the logic reads in our words, not theirs.

Give each part reader named `Element` and `Attribute` constants at the top for the spec identifiers it uses, then compare and index against those.
A spec string appears once, next to the code that uses it.
The same raw string that means two things gets two names, for example element `t` is `Element.Text` and attribute `t` is `Attribute.Type`.

Opaque code values, like the cell type codes, live behind a translation function that turns them into our domain types.
The codes never leave that one file.


## Tooling

When a Biome autofix contradicts a `tsc` error, disable the Biome rule, do not weaken the type checker.
The type strictness is chosen on purpose and is the stronger guarantee.
For example `useLiteralKeys` is off because it fights `noPropertyAccessFromIndexSignature`.


## Abstractions

A concept is first class.
It has an interface and named classes that implement it.
Never a bare function that returns an anonymous object.

The class owns every detail of its backing library.
All of fflate lives in `FflateZipArchive`, all of `saxes` lives in `SaxesXmlReader`.
Nothing else imports the library.

Picking the implementation is a separate free function, never a static method on the class.
`openZip` picks the zip implementation, `createXmlReader` picks the xml one.
Callers depend on the interface and the picker.
They never name the class.

Swapping the implementation should only touch the picker file (or at least very few other files).
If many callers have to change, it was not an abstraction.

Wire the default in one composition root that takes the abstraction in its constructor.
`Workbook` takes a `ZipArchive`, so a test builds `new Workbook(stubArchive)` with no real library involved.

Public exports are the interface and the picker, not the concrete class.

Dependencies point one way.
Low level leaves like `io` never import the concepts built on them.
A picker lives with its concept, not in a shared folder.

Not everything is a concept.
A pure utility with no alternative implementations stays a plain function.
`readAllBytes` is a function, not a class.
Do not wrap it in ceremony.


## Comments

Default to no comment.
The code and the names carry the meaning.

Never write a comment that states what something inherently is.

Never name the current implementation in the doc of an abstraction/interface.

Never justify the existence of an abstraction or a pattern.
If an interface needs a comment to defend why it exists, that is a design conversation, not a comment.

Do not restate the code in prose and do not teach the language or design patterns.

A comment is justified only when it records something the code cannot show: a non obvious external constraint, a spec or format quirk, the reason a surprising choice was deliberate.
If you cannot name that thing, there is no comment to write.

Doc comments on a type or method describe the contract a caller relies on, the inputs, the outputs, the errors, the edge cases.
They never describe why the abstraction exists or what backs it.


## Writing things down

Four prose files, each with one job. A fact belongs in exactly one of them.

`README.md` is for someone using the library.
What it does, how to call it, and the limits they will hit.
It says enough about why to help someone decide whether to use it, and no more.

`CHANGELOG.md` is what changed, in the reader's terms rather than ours.
Add to `Unreleased` as part of the change that earns it, never in a batch at release time.
By then nobody remembers which of a dozen commits a user would actually notice.
Describe the capability and the caveat that comes with it, not how it is built.
A refactor nobody can observe gets no entry.
A bug a user could have hit gets a `Fixed` entry even if nobody reported it.

`CONTEXT.md` is why.
The decision, what we turned down, the reason, and the questions still open.
This is the only place rationale lives.
When a decision changes, edit it in place and note what changed, rather than appending a second answer.

`MANUAL-CHECKS.md` is what no test can prove, and how to begin proving it by hand.
A check that cannot be acted on without asking someone is not written down properly yet.
