# Free slides

A pattern is a slide somebody designed: it has a place for a title, three cards and nothing else. A **free
slide** is a slide the software lays out itself, from a list of **blocks** in reading order: one text, three
cards and a chart, twelve charts, two hundred charts. It stands on the frame of one of the template's own
layouts (its background, logo, header texts and title) and its texts, panels and tables are copies of shapes
the template already has, so it looks like a slide of that template.

```ts
import { freeSlide, freeContent, isFreeSlide, renderDeck, validateDeck, FREE_PATTERN, FREE_BLOCK_KINDS } from 'slotdeck';
import type { DeckSpec, FreeBlockContent, FreeSlideContent, FreeBlockKind, FreeArrange } from 'slotdeck';
```

Free slides need nothing set up: a template can carry them when its manifest has a layout with a title that
leaves room under it, and running text in text boxes of its own (see "What a template gives a free slide").
They do not replace patterns. A designed layout is better wherever one fits; a free slide is for content that
no layout of the template shows.

## In one minute

```ts
const deck: DeckSpec = {
  version: 1,
  fields: { organizationDepartment: 'Riverbend Bike Share' },
  slides: [
    { pattern: 'cover', slots: { title: 'A year of rides', subtitle: 'Riverbend Bike Share, 2025' } },
    freeSlide({
      title: 'Rides by station',
      blocks: [
        { chart: north, caption: 'North Bridge' },
        { chart: tram, caption: 'Tram Square' },
        { chart: river, caption: 'River Gate' },
        { note: 'All three grew; Tram Square grew fastest.' },
      ],
    }),
  ],
};
const { pptx } = renderDeck(template, deck);
```

`freeSlide` writes the slide as a deck holds it. Rendering and checking need no option: a free slide is part of
the deck format.

## A free slide in deck JSON

```json
{
  "pattern": "free",
  "blocks": ["chart", "chart", "chart", "note"],
  "slots": {
    "title": "Rides by station",
    "b1Caption": "North Bridge",
    "b1Chart": { "type": "column", "categories": ["Spring", "Summer"], "series": [{ "name": "2025", "values": [7, 19] }] },
    "b2Caption": "Tram Square",
    "b2Chart": { "type": "column", "categories": ["Spring", "Summer"], "series": [{ "name": "2025", "values": [9, 24] }] },
    "b3Caption": "River Gate",
    "b3Chart": { "type": "column", "categories": ["Spring", "Summer"], "series": [{ "name": "2025", "values": [6, 15] }] },
    "b4Text": "All three grew; Tram Square grew fastest."
  }
}
```

| Field | Type | Meaning |
| --- | --- | --- |
| `pattern` | `"free"` | Always `free` (`FREE_PATTERN`). |
| `blocks` | array of kinds | What the slide is made of, in reading order. Any number, at least one. This field is what makes a slide a free slide (`isFreeSlide`). |
| `slots` | object | The content. The title is `title` and may be left out. The parts of block *n* (from 1) are `b`*n* followed by the part's name: `b1Text`, `b2Chart`, `b12Heading`. |
| `arrange` | object | Optional wishes about the arrangement: `columns` and `lead` (see "How the blocks are set out"). `dense` marks a long text a person asked to have whole, set by the planner (see "A long text a person asked for"). |
| `base` | string | Optional. The id of the pattern whose frame the slide stands on. Left out, the template's usual one. A pattern that cannot be a frame gives way to the usual one, and the deck is told (see "What the checks say"). |
| `colors` | `DeckColors` | Optional. Colours for this slide alone; see [colors.md](colors.md). |

### The kinds of block and their parts

A part marked `?` is optional: without a value it is simply not drawn, in `strict` layout as in `adaptive`.

| Kind | Parts (slot names for block 1) | Drawn as |
| --- | --- | --- |
| `text` | `b1Heading`?, `b1Text` (string or paragraphs) | Running text in the template's body style, with a heading over it when given |
| `list` | `b1Heading`?, `b1Items` (array of strings) | A bulleted list in the template's list style |
| `card` | `b1Icon`?, `b1Heading`, `b1Text` | A heading and a short text on the template's panel |
| `figure` | `b1Value`, `b1Label` | A big number with what it counts |
| `step` | `b1Heading`, `b1Text` | A card with a number the software writes: 1, 2, 3 in the order of the steps |
| `quote` | `b1Quote`, `b1Source`? | A quotation in the template's quotation style, and who said it |
| `note` | `b1Text` | One highlighted line across the slide: the bottom line |
| `chart` | `b1Caption`?, `b1Chart` (a `ChartValue`) | A native, editable chart |
| `table` | `b1Caption`?, `b1Table` (a `TableValue`) | A table in the template's table style |
| `photo` | `b1Photo` (an `ImageRef`), `b1Caption`? | A picture that fills its frame, cropped to it |
| `graphic` | `b1Graphic` (an `ImageRef`), `b1Caption`? | A picture shown whole: a diagram, a screenshot |

Give an optional part to every block of a kind on a slide, or to none: the room for a caption is kept for all
charts of a slide when one has a caption, so that they line up, and icons are drawn only when every card of
the slide has one.

A block without its required content is an error in a deck, and the planner leaves such a block out. A text
part takes a string, or a list of strings where it holds paragraphs: a number such as `87` is not text, and is
reported as `expected text`.

### Writing and changing one in code

`freeSlide(content)` takes the blocks with their content and returns the `SlideSpec`. `freeContent(slide)` is
its reverse, so that code adds, removes or reorders blocks without renumbering slot names by hand:

```ts
const content = freeContent(slide);                 // { title, blocks: [{ chart, caption }, ...], columns?, lead?, dense?, base?, colors? }
content.blocks.splice(1, 1);                        // drop the second block
content.blocks.push({ card: { heading: 'Next', text: 'Forty more bikes.', icon: 'bolt' } });
const changed = freeSlide(content);
```

```ts
type FreeBlockContent =
  | { text: string | string[]; heading?: string }
  | { list: string[]; heading?: string }
  | { card: { heading: string; text: string; icon?: ImageRef | string } }
  | { figure: { value: string; label: string } }
  | { step: { heading: string; text: string } }
  | { quote: string; source?: string }
  | { note: string }
  | { chart: ChartValue; caption?: string }
  | { table: TableValue; caption?: string }
  | { photo: ImageRef | string; caption?: string }
  | { graphic: ImageRef | string; caption?: string };

interface FreeSlideContent {
  title?: string;
  blocks: FreeBlockContent[];
  columns?: number;                           // how many blocks stand side by side on one line
  lead?: 'left' | 'right' | 'top' | 'none';   // where the first block stands when it is the main one
  dense?: boolean;                            // a long text a person asked to have whole (arrange.dense)
  base?: string;
  colors?: DeckColors;
}
```

A picture given as a string is its asset id.

## How the blocks are set out

The software decides, by rules. The same slide on the same template always gives the same file.

- **One kind of block** is set out as a grid: three cards side by side, six charts in two lines of three, two
  hundred charts in ten lines of twenty. The number on a line is the one that gives each block the shape it
  reads best in and uses the height of the room: twelve figures stand four by three, not six and six over an
  empty lower half. Up to eight cards or steps never stand more than four on a line, because a fifth makes
  every text small: five are three and two. Three charts on one line may be a little higher than wide, and end
  where the room ends. Of arrangements that suit alike, one whose charts each have a cell large enough for
  their names and numbers (about 17% of the slide's width and 7.5% of it in height) wins. Figures that stand by
  themselves may be three quarters as high as wide, so that three on a line are not lower than four in two
  lines.
- **A last line that is not full** lines up with the lines above it. Cards, steps and figures share the whole
  width when the line holds at least half as many as a full one (three cards over two wide ones), and stand in
  the middle when it holds fewer; steps keep their order, so the fourth and fifth of five follow the third.
  Charts and pictures are read as a raster: they keep the width of the lines above and start where those
  start, so that the fifth stands under the first. Running text, lists, quotations and tables share the whole
  width, however few they are.
- **Several kinds** are set out in bands, one under the other in the order given: three figures over two
  charts. A band of running text, a list or a quotation is as high as its lines need, with a line to spare,
  and what it does not use of its share goes to the bands that use all of theirs: a text of two lines over a
  row of cards leaves the cards the height it has no use for. A slide that mixes more kinds than bands can
  hold is set out as one grid of tiles, each on a panel.
  Tiles of one line are equally high, whatever they hold, and a last line of tiles that is not full shares the
  width. A table has no panel behind it: it brings its own ground. A chart has one only when the slide has
  blocks that are drawn on a panel; a slide of charts, tables and pictures alone has none.
- **One chart, table or picture among a few other blocks** leads: it takes about half the slide and the other
  blocks stand beside it, as a designed layout would do it. The blocks beside it end where it ends, and stand
  in full lines where they can: three figures beside a photo are a column, four are two by two.
- **A note** is a band across the whole width. A note given first stands over the other blocks; any other note
  closes the blocks over it: it stands one gap under the lowest of them, never under a hole that they leave, and
  under blocks that fill their room it stands at its foot. A note that fits one line is one line high: a written note is counted at half its size a character, the
  width of a note's heavy face, and its band keeps room enough that a second line would still end inside
  it. Notes are for the bottom line: beside other blocks they take two
  fifths of the room together, and more only where those blocks leave it. Notes that ask for more have their text set smaller, all
  at one size: the largest at which they fit. Smaller text takes fewer lines, so that size follows from how much
  the notes say, and short notes are never set smaller than long ones would be in their place. The gaps between
  them come down faster than the text. Such a slide counts as crowded. A slide of notes alone has the whole room
  for them.
- **In a deck read from the right** the first block of a line is at the right.
- **A card is as high as its text, and grows a little where there is room.** A line of cards or steps that is
  lower than its room is made higher before anything is moved: by up to 35% of its height, or 15% when the cards
  lie (they are 1.6 times as wide as high, or more), never higher than the template draws a card of that
  width, and only while what is written on the fullest card of the line, from the card's top to the foot of
  its text as PowerPoint most likely sets it, still fills 60% of the card. A card twice as high as its text
  reads as unfinished: past that the line settles in its room as other blocks do. Heading and text stay at the
  top of the card, as the template's own cards have them, and the card holds what it held: growing changes no
  size and no budget. Not beside or under a main block, where the cards share its height already, and not on
  a template that draws no panel behind a card.
- **Cards of one slide have one form.** Whether an icon or the number of a step stands beside the heading (in
  a card that lies) or over it is decided once for the slide, by its narrowest card: a short last line that is
  spread wider keeps the form of the cards above it.
- **Blocks that are still lower than their room** sit a little down the slide, not pressed to its top, with the
  notes that close them as one group. Blocks that leave three tenths of the room or more have two fifths of what
  they leave above them. Blocks that leave less sit down by evenly less, and blocks that leave a tenth of the
  room or less start at its top: there is no height at which a slide jumps from one to the other. What is left
  is under the note, never between the blocks and their note. A slide of nothing but running text, lists,
  quotations and tables hangs from the top line, as the template's own text layouts do.
- **A title of two lines** has the blocks start as far under its second line as they start under a title of
  one: the room starts lower by the line the title has more, and never less than half a gap under its last
  line. Its lines are counted at half its size a character. The title is still told when it is longer than a
  title may be.
- **The larger room** that a crowded slide may take is taken only when it gives some block more room. A slide
  is crowded where a block is under the least size its kind sits well in, where its notes had to be set
  smaller, where a chart (of up to 24) has a cell too small for its names and numbers or lower than 13% of the
  slide's width (eight charts over a bottom line are set out as ten are), where a card on a tile among other
  blocks needs a quarter more height than the tile has for what is written on it, and where card text longer
  than its budget had to be drawn smaller to stay inside its card. Blocks
  that are too narrow are no wider there, and such a slide is not moved up under its title for nothing. In
  it the height belongs to the blocks: two lines of blocks stand no further apart than 22% of a line's
  height, and a note stands three fifths of the usual gap from its neighbour.
- **Everything stays inside its block.** On a slide that holds too many blocks, a block can be lower than a
  line of its text or than its heading. Its shapes are cut to it, so that no shape of one block lies on a
  neighbour, and the texts that have no room are reported (see "What the checks say").

Two wishes can be stated in `arrange`:

| Field | Values | Meaning |
| --- | --- | --- |
| `columns` | whole number from 1 | How many blocks stand side by side on one line. |
| `lead` | `"left"`, `"right"`, `"top"`, `"none"` | The first block is the main one: larger, at that side of the slide or over the others. `none` sets all blocks alike. Left and right are sides of the slide, whatever the reading direction. |

**A wish the room cannot hold is not followed.** A block has a smallest size: the least width its kind sits
well in, and half the least height. The widths are 15% of the slide's width for a card or a step, 20% for a
text, 18% for a list, 30% for a quotation, 11% for a figure, 10.5% for a chart, 22% for a table and 9% for a
picture. Width and height are not alike because text is not: the size of a text follows the width of its
block, so twelve cards on one line have their text at about 5 pt, while a low block keeps the width of its lines
and seven cards one under the other still hold a heading and a text at 12 pt.

Where `columns` would leave some block under its smallest size, the slide takes the nearest number of blocks
side by side at which every block has it, fewer before more, and the deck is told (see "What the checks
say"). Twelve cards asked for on one line stand five on a line. Eight cards asked for one under the other
stand in two columns. A `lead` stays where it was asked for; the blocks beside it are relaxed among
themselves. When no number helps, the slide holds too many blocks for any arrangement: it is made as asked,
and counts as crowded.

### How large things are

- Texts take the template's own sizes. Like parts of like blocks share one size: every card text of a slide
  is the size the fullest card needs. Text that does not fit is set smaller, down to four fifths of its size;
  beyond that it is reported, like text over a slot's budget.
- **Spacing gives way before size (from 0.6.9).** Running text and lists that fit at the template's own spacing,
  and that closer spacing lets be set larger, are set larger. The space between their paragraphs closes up,
  to no less than 0.7 of what it is. Lines set wider than single close up by up to a tenth, never closer than
  single: 112% becomes 101% at the most, and a template at single gives none. It goes in steps, and the
  loosest step that fits is used. Like texts of a slide share one spacing. What a writer is told, and what is
  cut, reported or set whole, is counted at the template's own spacing, as before: closer spacing only sets
  a text that fits larger. In PowerPoint's own rendering, paragraphs closer than 0.7 of the usual space
  stopped reading as apart, and Hebrew lines stayed clear of each other at single spacing.
- **Card and step text is drawn where PowerPoint holds it.** Its words are also counted one after another as
  PowerPoint most likely sets them, each letter by its kind (a capital or an m is wider than an i), at about
  0.85 of a generous width a letter, which ordinary text exceeds on one line in twenty. Text whose words do not
  go into the lines of its box is drawn half a point smaller, or more, until they do. Text longer than its
  budget is drawn smaller still, down to three fifths of the card's size and never under 9 pt on a slide 13.33
  inches wide, and the slide takes the larger room where its cards are higher there; its budget stays the one
  stated at four fifths, so it is still reported. Only where it would leave its card even then does the
  heading give way: it is set smaller until the text stays inside, or at the largest size at which it takes a
  line fewer. The box of a card text is drawn as high as its last line may hang into the card's padding, so
  that what PowerPoint sets stays inside its own box.
- A line of text is counted as high as the template sets it: 1.2 times its size, or more where the shape a
  text is a copy of has a wider line spacing. With a line spacing of 130% a line is 1.56 times its size high,
  or up to a tenth less in running text and lists whose spacing gives way (see below).
  The last line of a text may hang a quarter of its size under its box, the space under its letters, and no
  more than 3 pt on a slide 13.33 inches wide.
- How long a line is can only be estimated without the template's fonts: a character is taken to be 0.55 of
  its size wide, and a line that breaks at words 85% full. The estimate is generous on purpose, and every
  budget is stated with it. A template that states how wide its text is set (`characterWidth`, from 0.6.7)
  has its running text, texts and lists, counted with that width instead; see "How wide a template's text is
  set" in [text-fit.md](text-fit.md). On a template with wide line spacing the text of a card or a step is counted
  twice, each time with the length and the height of a line taken the same way: by the estimate, with lines
  of the usual height, and as PowerPoint sets ordinary text (a character 0.42 of its size wide), with lines as
  high as the template sets them. It has to pass both. So card text on such a template is as large as on any
  other, a card is as high as the lines its text takes, and a box still holds no more than its real lines do.
  Before there is content the budget of a slot is counted with the estimate at the real height of the lines:
  a writer is told what the box holds for sure.
- A card or a step too low for a heading and a line of text at those sizes has both set smaller until they end
  inside it: six cards one under the other have smaller text than three side by side. A card that is too low
  even at the smallest size holds no text: what is written for it is reported on its slot. Short card text is
  set larger where its card has the room, but never as large as the heading over it, and never larger than the
  running text or the list of its slide, which it would then outweigh.
- The heading of a card or a step has one line, or two when it is longer than a line. It has the size the
  card's width and height give it, and keeps it when the text under it is set smaller: a heading is not set
  smaller because its text is long. Every card of the slide
  keeps the lines of the longest heading, so that the texts start on one line. Headings that take two lines
  are set up to 15% smaller when they then take one, so that no card keeps a line empty over its text for one
  heading a few characters too long. A heading whose words do not go into two lines is set up to 15% smaller
  when it then takes two, and otherwise has its three lines; it is still reported as longer than a heading
  may be, but it is not drawn over its text.
- A text that can take several lines (running text, a list, a quotation, a note, the label of a figure, a
  heading of two lines) is written with its lines at least as far apart as its size asks, whatever the line
  spacing of the shape it is a copy of. A template may set a heading of one line at 70%; two lines set so
  run into each other. A text of one line keeps the spacing of its source, and with it its place.
- A list is fitted point by point: every point takes whole lines in the width its bullet leaves, and the
  points stand half their size apart, or closer where spacing gives way. Six points of a line and a half
  each take twelve lines, not the nine their length alone comes to, and the list is set smaller, or
  reported, by that count.
- A **figure** is as large as its card and the number of figures allow. Six figures or more have the
  template's own size for a big number; fewer have more, by the square root of six over their number, up to
  2.4 times that size and never over two and a half times the title: on the demo template one figure is set
  at 85 pt, two at 69, three at 56.5, four at 48.5 and six at 40. The value does not break, so it is set as
  small as the longest one needs to stay in its box and to leave its label the lines PowerPoint most likely sets
  it in. It is never smaller than the template's own size for a big number, or 0.42 of its card's height where
  that is less, in a card as wide as the value needs: so fewer figures in a room are never smaller than more,
  and a long label does not make its value small. Its label is about half its size, never under the size of
  small running text where the card has the room, and has one to three lines; where those lines do not fit
  under the value, the label is set smaller. Value and label sit together in the middle of the card, with as
  much space over the value as under the label, and cards of one height have them on one line. A figure that is the slide's only block has no panel: it stands
  large where the title starts, a little above the middle of the room. A value longer than a figure may be
  (about nine characters where three figures stand side by side) is reported on its slot.
- The **source of a quotation** is the small line under it: half the quotation's size, never under a caption's
  and never over running text's, in the template's accent (the colour of its big numbers) where the
  quotation has another colour. A quotation has its text at the foot of its box, so its source follows its
  last line at half its size, however many lines PowerPoint sets it in. Quotations side by side have boxes of
  one height, and their sources stand on one line, each as far under its own last line.
- One sentence alone on a slide is set large. A long text alone is set in two columns of the same height. A
  text alone keeps a line a person reads in one go, 66 characters at the size it is set at; under a heading it
  is set smaller, and its line is shorter with it.
- A deck that asks for larger text (`textScale`) never gets smaller text for it. Larger text takes room from
  its neighbours. On a slide where some text would then have to be set smaller than it is without the wish, the
  slide takes half of the enlargement, or none, and the deck is told which (see "What the checks say"). A text
  that would only lose half a point, to rounding or to a block that got a little narrower, is held at the size
  it has without the wish, when that leaves no text with less room than it had, and the slide keeps its larger
  text. Text that is no smaller and does not fit at the larger size is reported on its slot, like any other.
  Where the manifest lets the renderer set text smaller to fit (`shrinkText`, on a template loaded with fonts),
  a text of a free slide is set no smaller than it is drawn without the wish: it is held at that size, and
  reported when it still does not fit. From 0.6.7 the size a text may not go under is counted as 0.6.6 counted
  it, its headings on lines of their own and a character no narrower than the estimate: a text that the count
  of headings run in, or of the template's width, sets larger without the wish does not cost the rest of the
  slide what 0.6.6 gave it.
- A slide **crowded** beyond what reads well is made all the same, in a larger part of the slide, and the deck
  is told (see "What the checks say").
- A **chart** is drawn in as much detail as its box reads at. Its slot states how many categories and series
  read well at that size (`maxCategories`, `maxSeries`). These are soft limits, as on the chart slot of any
  layout: a chart over them is drawn, and reported on its slot.

  | Detail | When the box is | What is drawn | Categories and series that read well |
  | --- | --- | --- | --- |
  | `full` | at least about a quarter of the slide's width and of its height (each of six charts in two lines still is) | Axes, gridlines, numbers at the bars where the slide has them, a legend | 8 and 3 |
  | `compact` | smaller, down to a strip of about 45 pt, where its text is still 7 pt or more | The names of the categories, and the numbers at the bars where the slide has them or else a value axis of three to five labels; a legend; no gridlines | 3 to 8, by its width, and 2 |
  | `spark` | lower or narrower than that, or where its text would be under 7 pt, or one of more than 24 charts on the slide, or a low chart on a tile among other blocks | The bars or the line alone, filling the box | 12 and 1 |

  Sizes in points are for a slide 13.33 inches wide; a narrower slide takes less. No text of a chart is set
  under 7 pt, under which nobody reads a projected slide: a chart whose place carries no more is drawn bare,
  its bars to the largest value and its caption over it, and the deck is told that the charts of that slide
  are drawn without axes, legend or labels. A slide whose charts would be bare in the usual room takes the
  larger one where that keeps them readable: on the demo template up to fifteen charts over a bottom line
  keep their names and numbers (twelve stand four by three), and sixteen or more are bare. The detail
  follows from the place alone, so a slot says the same of its chart before and after the chart is written.
- **Numbers at the bars in place of a value axis is decided once for the slide.** Every small bar and column
  chart has its numbers at the bars where each of them has room for them and no small line chart stands among
  them (a line carries no numbers), and none has them otherwise: one slide never shows small charts with
  numbers and no axis beside small charts with an axis and no numbers. A small chart gives the axis up for its
  numbers; one with a value under nothing keeps its axis. A full chart always keeps its axis, so numbers are an
  addition to it: like full charts have them where every one of them has room for them, whatever an unlike
  chart beside them can hold (a bar chart of shares keeps its numbers beside a column chart of millions that
  has no room for its own).
- **The charts of a slide drawn in one detail set their text at one size**, the smallest any of them needs: a
  bar chart that names many categories beside column charts has its names no smaller than theirs.
- **Like charts share one scale.** Charts of one type (not pies), one format and one number of series, none
  with a value under nothing, are a group when the largest value of one is no more than five times the largest
  of another. A group has one value axis, so that a bar of one chart can be read against a bar of the next;
  further apart nothing is forced, and each chart has its own. Where the axis is shown it has three to five
  labels at round values, fewer where its plot is short: a step of 1, 2 or 5 times a power of ten, or 25 times
  one, and whole steps before steps that write a decimal on every label (0, 2, 4, 6 and not 0.0, 2.5, 5.0; 0, 1M,
  2M and not 0.25M, 0.50M). Of the steps that serve, the one whose axis the values fill best is taken, then the
  one with fewer labels. The axis ends at the half step over the largest value, which may be over its last
  label: a bar of 105 stands on an axis labelled 0, 50, 100 that ends at 125, not on one of 0, 100, 200. Where
  that would take a label too many, it ends just over the largest value. Its labels are written with the
  decimals of the step, not of the values (0, 1, 2 and -10%, not 0.0 and -10.0%). Bars start at nothing. A line
  starts there too, unless it stays far above nothing (its lowest value at 60% of its highest or more): it then
  starts at the round step under its lowest. A chart with a value under nothing starts at the whole step under
  its lowest value; where numbers stand at its bars, a step lower only where the room under the lowest bar is
  less than its number takes, so that the number stands clear of the names. Where the numbers stand at the bars
  instead of an axis, and in a bare chart, a bar chart ends at the largest value itself: the longest bar of the
  group uses the whole of its plot. A group also shares the way its numbers are written.
- **One legend for a slide of like charts.** When every block of a slide besides its notes is a chart, none of
  them a pie, and all show the same two or more series, the slide says once what the colours stand for: a key
  and a name for each series on a strip at the top of the room, from the edge reading starts at, and no legend
  in any chart. Every plot has the room its own legend took. Among other blocks, or where the series differ,
  each chart keeps its legend.
- **What a chart is drawn with**, where its place is known: a plot whose place is counted from its text size,
  so that names, axis and legend take what they need and no more; names of categories that stand upright,
  every one of them beside horizontal bars and every so many under columns; under a plot with a value under
  nothing they stand at its foot. Under columns a name has the width of its place in the plot, and every so
  many are written where that is too narrow for each. Beside horizontal bars the names have their own width
  and nearly two sizes more: with less, PowerPoint breaks `Q1` into `Q` over `1`. Over columns a number at the
  bars takes 1.3 times the text size, and the names under columns and lines 2.1 times, or 1.7 times over a
  legend of the chart's own, whose line gives them air. Its text follows its frame: a full
  chart has the template's own size and no more than 6% of its frame's width and a tenth of its height, a
  small one 7 to 10 pt and no more than an eighth of its frame's height, so that names and numbers leave the
  bars most of a low chart. Numbers are written with thousands apart, with the decimals the values have (46.5%
  is not shown as 47%). A million and more is written in millions, with the decimals the values have there up
  to two, and otherwise two as a reader says them (`1.25M`, and `1,385,000` as `1.39M`), or three where two
  would write two values alike; thousands are written so only where that loses nothing (`250K`). A bar is never thicker than it reads: a
  column takes at most 11% of the plot's width. Slices and series past the end of the template's chart colours
  are tints of those colours, never the first colour again.
- A **caption** of a chart, a table or a picture has one line. One that takes two is set smaller, by up to
  15%, where it then takes one; otherwise its strip is as high as its lines, up to three, and what it names
  starts under it, as it is really set. Its lines are counted as PowerPoint most likely sets its words, each
  letter by its kind. Over a chart the strip is 1.25 times its size high and the chart starts a fifth of its
  size under it: the plot keeps space over its bars of its own.

- A **table** is as high as its rows need, and a row as high as the lines of its fullest cell as PowerPoint
  most likely sets them (its words one after another, each letter by its kind, the heading row in its heavier
  face): rows set in as many lines are equally high, and a row PowerPoint sets on one line is not as high as one
  of two. Beside other blocks its rows take up to two fifths more height where its block has it, so that a
  table of a few rows does not hang at the top of a block whose neighbour fills it; a table alone hangs from
  the top line at its rows' height. Its columns are as wide as what is written in them asks: a column of long
  names is wider than a column of short numbers, so that a long cell takes one line where its column has the
  room. Where the table has not the width for that, every column is first as wide as its longest word, its
  letters counted by their kind (`Medium` is wider than `little`), with what the template's table keeps free at
  the sides of a cell, and a number with the short unit after it counted as one word (`2.4 minutes`); the width
  that is left is shared by what each column asks for beyond that: no word is broken, and no number parted
  from its unit, where another column has width to give. A column then gives up to a fifth of its width to
  another where that sets some row in fewer lines and no row in more. Where the longest words alone are
  wider than the table, its text is set smaller, by up to a fifth, until they fit. A column of numbers, one
  whose every cell starts with a number (`4,120`, `-3%`, `6 days`), ends on one edge, its heading with it, so
  that a number does not run into the start of the next column; every cell has its text in the middle of its
  height. With many rows, or
  with cells of several lines, its text is set smaller so that it fits: down to 6 pt on a slide 13.33 inches wide. A table that does not fit even then is drawn as
  asked and reported on its slot. So is a table whose cells take its text under four fifths of the size it
  would have with a word in every cell: it is still drawn inside its box, and shorter cells make its text larger.
- A **photo** fills its box by being cut, up to a box 3.2 times as wide as high; a wider one, or a very high
  one, is not filled edge to edge, and the photo then stands at the edge reading starts at. Its caption
  starts under its corner and is as wide as it is. A graphic is shown whole in the middle of its box, and its
  caption stands under its middle.

### What a slot says before the slide is written

A slide that is planned has no content yet, and its slots have to say how much each will hold. The slide is
then set out as it will be when every text is about as long as texts of its kind usually are: a card or a
step text of about 70 characters, a quotation of about 100, a list of about four points of 35, running text
of about 300, a note of 110. The sizes are found by the same fitting that content gets, and each slot states
what its box holds at the size that ends on.

- **A budget is what the slide really takes.** A card is planned as high as a text of 70 characters needs,
  not as high as its room. Nine cards on the demo template are told 54 characters, twelve 42 and sixteen 22:
  what the slide that is laid out holds at the size its text is then set at. Text written to its budget, or
  to nine tenths of it, is not reported on its slot.
- **Nobody is told to write one character.** A box that holds no line of its text even at the smallest size
  states one short line's worth, what its width holds at that size and never under 12 characters. With
  content such a box keeps the honest report: `text is 38 characters; the slot fits about 1`.
- **A slot says its limits once, in its fields**: `maxChars` and `maxParagraphs` for a text, `maxCategories`
  and `maxSeries` for a chart, `maxColumns`, `maxRows` and `maxCellChars` for a table. Its `description` says
  what the part is for and, for a text with room for it, how much to write (`Write 40 to 70 characters`). It
  repeats no limit and has no full stop at its end, so that whoever shows it to a writer can put the limits
  after it. A list whose slot holds one paragraph says `One short point`, not `One short point in each item`.
- **A note is told no more than a note alone.** Many notes are set smaller, and a line of small text holds
  more; the budget stays that of the one line to take away.
- **Optional parts are counted in by rule**: an icon for every card when an icon set is at hand, a heading
  over each text of a slide of several blocks, a caption for each chart, table, photo or graphic of a slide
  that has two or more of the kind, the source of a quotation. A change to a deck names the optional slots
  that hold a value on the slide (`withParts`, such as `b2Caption`), and those are counted in besides, so
  that a caption or a heading the slide has can be written again.
- **A text a person asked a length for is planned at that length** (from 0.6.1). Where a plan has `words` for
  the slide, its running texts are set out for that many characters (about 6 to a word, 6.5 in Hebrew) in
  place of the usual 300, and their budgets say what their boxes then hold. On the demo template a text of 300
  characters alone is set at 22 pt on a short line and told `Write 290 to 550 characters`; one of 1,200 is set
  at 15.5 pt in two columns and holds 1,377. The writer is told the length asked
  (`Write 200 words, as was asked: 1200 to 1370 characters in all`). See "Asking for an amount of text" in
  [generating.md](generating.md).

### A long text a person asked for

From 0.6.6 the headings of a person's document in such a text are set in bold, each at the head of the
paragraph under it, an em space between them (`SlideSpec.headings`, see [deck-format.md](deck-format.md)). A
heading set so shares a line with its text, so no heading ends a column with its text in the next. The heading
a slide's text opens under is its title. From 0.6.7 the count that sizes the text counts a heading set run in
as it is set:
- it takes no line of its own;
- its bold letters and the em space go to the head of its paragraph, and its space after goes before it.

The count asks of each heading the renderer's own question, for a line a character shorter. A heading over a
list item, a table row or another heading, or too long for a line, is counted on a line of its own, as it is
set. While a deck is planned, a document's text that code sets as it is carries its headings, so the plan
counts it as the slide will be set, in either room. Whether a text needs the larger room is still decided as
0.6.6 counted it, its headings on lines of their own. Counted run in, a text of short parts fits the first room
more often, and kept there it came out smaller than in the larger room, where it is set in columns.

A person may ask for an amount of text ("500 words on each slide") or give their own text for the slides. From
0.6.2 such a text stays on the slide it was asked for, set smaller where it must. The switch is
`arrange.dense`: a free slide of one text, with or without a heading, that carries a text someone asked to
have whole.

```json
{
  "pattern": "free",
  "blocks": ["text"],
  "arrange": { "dense": true },
  "slots": { "title": "Our plan for 2027", "b1Text": ["The first of up to twelve paragraphs.", "The second."] }
}
```

- **The planner sets it.** Every free slide of one text that the planner gives an asked length (`words`) or
  passages of a document is marked dense, in the plan and in the deck. No model sets it, and no prompt shows
  it. A change to a plan or a deck keeps it. Code that writes a deck may set it too.
- **A text that fits is set as large as it fits.** As without the flag, down to four fifths of the
  template's size, the most a free text is set down to without it. The flag sets smaller only a text alone that
  does not fit there: before 0.6.3 one of more than 600 characters, from 0.6.3 also a short one of many
  paragraphs, such as a list of fifteen short points. It is there so that the paragraphs a writer chooses
  never make the text run over its box: where they need more room than the plan counted, the text is set
  smaller instead. Every other slide with the flag is laid out as if it had none.
- **Down to the least size that reads.** Such a text is set as large as it fits, half a point at a time, down
  to 9 pt on a slide 13.33 inches wide, and less on a narrower one by its width, never under 5 pt: 7 pt on a
  slide 10 inches wide. That is the size under which the checks say a slide's text is small. From 0.6.7 a
  template can set that size itself (`leastTextSize` in its manifest, see
  [manifest-format.md](manifest-format.md)). Such a text is then set no smaller, also where the first room
  would hold it at four fifths of the template's size, and is cut where it needs smaller. A value over the
  template's own size of running text is taken as that size. A lower value also lowers the size under which
  the checks say a slide's text is small.
- **In the larger room, in columns.** It takes the larger part of the slide that a crowded slide takes, under
  the title. It is set in two columns, or three where two would make its lines longer than about 95 characters
  as PowerPoint sets them. From 0.6.7 it also takes three where two do not hold it at a size, as long as a line
  in three still holds 45 characters. Before, it went down half a point at a time in two columns, then jumped to
  three with the last column part empty. Where it fits at no size, it is set at the least size, in the
  columns its lines take there, as before.
- **Counted as PowerPoint sets it.** A character of such a text is counted at 0.48 of its size wide, where other
  text is counted at a generous 0.55. In PowerPoint's own rendering of 3,035 lines of such text at 6 to 17.5 pt,
  in Hebrew and in English, a character took 0.41 of its size on a middling line, 0.44 on one line in twenty,
  and never more than 0.47. Lines still break at words. A template that states how wide its own text is set
  (`characterWidth`, from 0.6.7) has its running text counted with that width instead, set dense or not: see
  "How wide a template's text is set" in [text-fit.md](text-fit.md).
- **One rule says whether it fits** (from 0.6.3). Whether such a text fits, and how much it holds, is said by
  the same count when the slide is planned, when its writer is told a length, when code sets a document's text
  as it is, and when the deck is checked. It counts by length: each paragraph takes whole lines, as many as
  its length takes with room lost where lines break at words, a short one too (twenty paragraphs in capitals,
  each a little shorter than a line, took two lines each: counted one line each, they ran 150 pt past their
  box); the lines go into the columns one after another as PowerPoint fills them, and each column holds its
  lines and the space after each of its paragraphs in its height less one whole line. PowerPoint takes the
  line that ends a paragraph into a column only where the space after it fits too: in its own rendering of 96
  decks of such text, all 215 column breaks fell so. The line each column keeps free is there because
  PowerPoint does not always set a column as it is counted: a text of two columns that fitted came out now and
  then with its first column a line longer, and words of wide letters take more of a line than their length
  says. No measure of the words, of the kind of their letters or of a language counts: the face a template
  sets them in is not known.
- **Planned for what the writer may write.** While the slide is planned, its text is counted as the most the
  writer is told: 1.15 times the length asked, in up to twelve paragraphs, or a document's text in as many
  paragraphs as it has, its headings among them. Those paragraphs may be of any length, and each may end a
  character into a line, so each is counted to lose up to a line, with the space after it. The budget is what
  the slide holds at the size chosen however the paragraphs fall, and the slot takes that many paragraphs. A
  writer inside both is not sent back. A text of an asked length that the slide holds at four fifths of the
  template's size is set as without the flag instead, smaller where what is written needs it.
- **A cut text is told twelve paragraphs at most.** Each paragraph a writer is told costs room, so where the
  slide does not hold a document's text, its writer is told twelve paragraphs at most and joins the rest:
  told the 21 paragraphs of a document, its headings among them, a slide holds about a fifth less. Where the
  slide holds all of the text's characters in more than twelve, its writer is told the most that hold them,
  and nothing is cut. Either way it is told no more than leave half the room of its columns to its lines.
- **A document's text that code sets is judged as it is.** Where code would set a document's text word for
  word (see "Making a deck from a document" in [generating.md](generating.md)), the slide is planned for its
  own paragraphs, with the headings code sets between them. Code sets it only where the slide holds it so;
  otherwise its writer writes it, told what the slide holds.
- **Beyond the least size it is cut.** A text the slide does not hold even at that size is planned at it, its
  budget is what the slide holds there, and the plan says what is cut before anything is written
  (`planTextFit`).
- **Once written, the text itself is judged.** A text that does not fit is always told so: what the box holds
  in as many paragraphs as it has, however they fall, or, where it has more than leave half the room of its
  columns to its lines, that it has too many. It is never told to join them into twelve. Its box is as high as
  its columns are where they come out alike, and the line they keep free. PowerPoint may fill them so that the
  last column comes out the shorter, at times almost empty.
- **Code that changes the blocks keeps it.** `freeContent` gives `dense: true` for such a slide and `freeSlide`
  writes it back (`FreeSlideContent.dense`), so `freeSlide(change(freeContent(slide)))` keeps the flag.

What it gives, from the plan and from PowerPoint:

| Template | 500 words | Set at | The most one slide holds |
| --- | --- | --- | --- |
| The demo template, 13.33 inches wide, English | about 3,000 characters | 10 to 10.5 pt, in 3 columns | about 4,350 characters, some 720 words, at 9 pt |
| A right-to-left template 10 inches wide, Hebrew | about 3,250 characters | 7 to 7.5 pt, in 3 columns, cut to about 2,900 | about 2,900 characters, some 440 words, at 7 pt |

A heading over the text takes some of that room: under one, the right-to-left template holds about 2,480
characters at 7 pt. A text written in more paragraphs holds less, and a cut text is told twelve at most. The
rule was measured in PowerPoint (0.6.3) on 666 planned slides of both templates, in their own faces:
documents, lists and lengths asked, set by code, or written at the top of what the writer was told in some 30
shapes. Every line ended inside its box, and no writer inside what it was told was sent back. Of 693 slides set
straight from a deck, at lengths over and under what they hold, every one that ran past its box was told so.
In 3 to 100 paragraphs alike the rule accepts no longer text than 0.6.2 did. It knows no face: at the longest
it accepts, capitals and long unbroken words such as web addresses ran up to 12 pt past the box unsaid in the
templates' own faces, and up to 223 pt in a wide face. In a wide face a text written at the top of its told
range can run past too: 500 German words on the demo template ran 24 pt past. A list of short items, each a
little shorter than a line, is counted two lines an item, so a slide can cut it, or hold it only with its
items joined, where it would fit whole; 0.6.2 set such lists whole, and most of them ran past their box. A
slide that holds a document's text only with its paragraphs joined says so: `the text from the document for
this slide (p1-p27) has 53 paragraphs, and the slide holds it only as 25: its paragraphs are joined`.

A list of many short points is such a text too (from 0.6.3): where one column of it does not hold its
paragraphs at four fifths of the template's size, it is set in two columns in the larger room. In 0.6.2 a list
of fifteen points of 30 characters set dense was told six points in one column, and cut by about a sixth.

A list, a chart, a table or cards beside the text keep the slide from being set dense: such a slide is set out
as without the flag, and holds what it holds. So does a second text of a deck's slide (`planEdit`). A plan's
slide of several texts, where nothing is written yet, becomes one text set dense where its texts would be cut.
From 0.6.3 the person's text comes first in a plan (`planOutline`, `editPlan`): a slide that carries the text
of a document and would be cut beside its other parts is its text alone, set dense, where that keeps more of
it, and so is a slide whose other parts leave it under half of an amount of text asked for it. In a change to a
plan that is so only for the parts the plan had: a part a change asks for stays, in that change and the ones
after it. With free slides off no slide is set dense.

## What a template gives a free slide

Nothing has to be set up. When a template is loaded, the software reads from the slides its patterns stand on:

| What | Where it comes from |
| --- | --- |
| The frame | The layout whose title sits where most of the template's titles sit, when it leaves room under that title. Its background, logo, header and footer texts, page number and title stay; its own content is taken off. |
| Text styles | The shapes the template's patterns use for running text, headings, big numbers, captions, lists and quotations. Each text of a free slide is a copy of one of them: typeface, colour, weight and line spacing come with it, the line spacing closed up a little where running text gives way. The space between paragraphs is the software's own. |
| The card | A tinted or outlined box that the template sets text on. |
| The note | A text the template sets on a box of its own (a takeaway line), or a strongly filled box with the text style used on it. It is set no larger than nine tenths of a heading: the bottom line is not the loudest text of the slide. |
| The quotation | The template's own quotation. A template without one has it made from its heading: light, in the colour of running text, slanted when the template is read from the left and upright when it is read from the right, whose letters have no slant. |
| The table and the picture | A table of the template, with its fills and borders; a picture of the template, for its rounded corners. |
| Chart colours | The series colours of the template's chart slots, or else its main colours. |
| Margins and spacing | Where the template's own layouts keep their content, and the gap they leave between neighbours. The room starts where the first thing under the layouts' title starts. Where the frame of a free slide draws something on the line a subtitle takes in those layouts (a slot named `subtitle` or `subheading`), such as a text of the deck that stays on every slide, that line is the title's: blocks start under it, and the room ends lower by what that takes, as far as the layouts' content goes. A free slide has no subtitle of its own, so where nothing stays on that line its blocks start on it. Where the content area leaves clearly more margin at the side reading ends than at the side it starts (a hundredth of the slide's width or more), the room is made wider at that side, as far as the frame above the title reaches. |

A template without a card box draws cards as headed texts, under a thin line when the template has one. A
template without a table has one drawn in its main colour.

**A frame leaves room.** A layout is a frame only when blocks get at least 45% of the slide's height under its
title, measured from where they start to the bottom of the content area. Layouts made for content leave more
than half the slide. A cover, a section divider and a closing slide leave 40% at most, and are passed over. A
template that has no layout with that much room is stood on the one that leaves most. On such a template a
pattern that is named and leaves less is told by comparison with that one (`... and "section" leaves them
40%`), and the one that leaves most can be named as the frame.

**A title in a placeholder stays where it is.** The usual title of a real template takes its position from the
slide layout, or from the master. It is read from there and filled in place; it is never moved or made wider.

**On a dark frame**, the text and the lines of a chart, and the text of a table that the software draws, are
set in the colour of the slide's own text, which reads on that frame. On the template's usual frame a chart
keeps the template's look.

Where the software's choice does not suit, the manifest states it:

```json
{
  "free": {
    "base": "text-only",
    "styles": { "heading": { "slide": 7, "shape": "7" }, "panel": { "slide": 5, "shape": "7" } }
  }
}
```

| Field | Meaning |
| --- | --- |
| `free.base` | The id of the pattern whose frame free slides stand on. It needs a title that leaves room under it (see "A frame leaves room"), and must not be a pattern with a row of items. |
| `free.styles` | For a part of a free slide (`body`, `heading`, `figure`, `caption`, `list`, `quote`, `note`, `panel`), the slide number and shape id of the template shape it is a copy of. |
| `"free": false` | Free slides are not offered when a deck is planned. A stored deck that has one still renders. |

A studio sets these with `setManifest`, which refuses a frame or a shape that cannot be used and says why.
`checkManifest` reports a setting that points at nothing, and one that is passed over when a slide is made: a
frame without room under its title, a style that no layout of the manifest uses as a text any more. See
[template-studio.md](template-studio.md) and [manifest-format.md](manifest-format.md).

A manifest that `changeManifest` gave back takes effect at once, also on the package that is already open
(`{ pkg: template.pkg, manifest: result.manifest }`). A manifest object that is edited in place is not seen:
make the change through `changeManifest`, or load the template again.

## What the checks say

`validateDeck` and `renderDeck` check a free slide like any other. Pass the loaded template to `validateDeck`
so that the slide is laid out and its text is checked against the room it gets:

```ts
const issues = validateDeck(template, deck, { assets: ['pier', 'clock'] });
```

Given only the manifest, `validateDeck` checks what needs no template: that every block is of a known kind, has
its required parts and holds values of the right kind. It says nothing about fit.

| Message | Severity | Meaning |
| --- | --- | --- |
| `a free slide needs at least one block, such as "blocks": ["text"]` | error | `blocks` is empty. |
| `"blocks" names "poem"; a free slide is made of text, list, card, ...` | error | A kind that does not exist. |
| `this slot is required` on `b2Text` | error | A required part of block 2 has no value. |
| `expected text` on `b2Value` | error | A text part holds a number, an object, or a list with one of them in it. |
| `a free slide cannot be made from this template: ...` | error | The template has no layout a free slide can stand on. The sentence says what is missing. |
| `text is 310 characters; the slot fits about 120` on `b1Text` | warning | The text of that part is longer than its box holds at the smallest size it may be set in. |
| `text is 38 characters; the slot fits about 1` on `b1Text` | warning | The box of that part is lower than one line of its text, even at the smallest size: a card among too many, one under the other. The text would be drawn across the card's edge. |
| `slide 4 holds 60 blocks, as asked; its smallest text is set at 5.5 pt` | warning, on the deck | The slide was made as asked, and some of its text is under 9 pt, on a slide 13.33 inches wide (in proportion on a narrower one), or under the template's own `leastTextSize`: smaller than anybody reads on a projected slide. |
| `slide 4: 12 side by side were asked for; the slide holds 5` | warning, on the deck | `arrange.columns` would have left some block under its smallest size. The slide was set out with the nearest number that does not (see "How the blocks are set out"). |
| `the chart has 10 categories; the slot fits 8` on `b1Chart` | warning | The chart has more categories, or more series, than read well at the size it is drawn at. It is drawn all the same. |
| `slide 4 holds 200 blocks, as asked; 200 charts are drawn small, without axes, legend or labels` | warning, on the deck | The charts are too small for anything but their bars or lines. |
| `slide 4 holds 9 blocks, as asked; the deck asks for larger text, and this slide has no room for it` | warning, on the deck | The deck's `textScale` is above 1, and at that size some text of the slide would be smaller than without it. The slide is made as a deck without the wish would have it. Also `... has room for half of that`. |
| `the table has 60 rows; the slot fits 32` on `b1Table` | warning | The table has more rows than its box holds at the smallest text size. PowerPoint will draw it past its box. |
| `a cell is 60 characters; cells fit about 24` on `b1Table` | warning | A cell takes more lines than its row has with the text at four fifths of the size the table would have with a word in every cell. Down to the smallest size the table is still drawn inside its box, in smaller text; beyond that PowerPoint draws it past its box. |
| `slide 4 holds 9 blocks, as asked; one of them reaches below the room the slide has, and may be cut off at its edge` | warning, on the deck | A block was set out below the room for content. The software lays nothing out there; this is said should it happen all the same. |
| `slide 4 was planned on the frame of the layout "x", which the template no longer offers; it stands on "y"` | warning, on the deck | `base` names a pattern that is gone. |
| `slide 4 was planned on the frame of the layout "x", which cannot carry a free slide (...); it stands on "y"` | warning, on the deck | `base` names a pattern that is no frame. The brackets say why: its title sits low on the slide, it has a row of items, it has no title. For a low title they give the room it leaves, and what a frame leaves: 45% or more, or, on a template with no such layout, what the layout stood on leaves. |
| `pattern "free" has no such slot; ignored` on `b9Text` | warning | A slot of a block the slide does not have. |

The warnings that begin `slide 4 holds`, `slide 4:` or `slide 4 was planned` carry no slide number in
`Issue.slide`: they are notes on what was asked for, not requests to write the slide again. A size in such a note is given in half
points, as it is written to the file. Where the small text is a table's and its cells are what makes it small,
the table's slot is told as well, with the slide's number, and shorter cells change it.

A table's slot states what its room holds (`maxRows`, `maxColumns`, `maxCellChars`). Once the table has its
content these are worked out for it, with the text at four fifths of the size the table would have with one
word in every cell, and never under the smallest size. A deck's larger text is left out of that count. A table
that fits at that size or a larger one is never reported, whatever lines its cells take. One that needs
smaller text always is.

Characters that no file can hold (control characters such as a vertical tab, which comes with text copied out
of a text box) are left out of every text. A vertical tab breaks the line as a newline does.

The other functions that read a deck know free slides: `extractAssets` finds their pictures, `checkDesign`
looks at their text for overflow and lone last words, and `slideAlternatives` returns no alternatives for one,
since a free slide is its own layout.

## Two hundred charts

A slide may hold any number of blocks. Two hundred native charts on one slide make a file of about 700 KB,
built in a quarter of a second, and PowerPoint opens it in a few seconds. Each chart is a chart of its own with
its data in an embedded workbook, so a person can click one and edit it. At that size a chart is a sparkline:
there is no room for an axis. Like charts among them share one scale, and the longest bar of them all touches
the top of its frame.

## Free slides when a deck is planned

The planner offers free slides by itself: `planOutline`, `planDeck`, `generateDeck` and the chat tools need
nothing set. The full account is under "Free slides" in [generating.md](generating.md); this is what it comes to.

**When one is offered.** The list of patterns a model plans from ends with a line for `free`, when:

- the option `free` is not `false` (`PlanOptions`, `GenerateOptions`, the chat tools);
- the template's manifest does not say `"free": false`, and the template can carry a free slide: it has a layout
  to stand one on, and the brief gives the deck fields that layout shows;
- no list of patterns is in force, or the list names `free`. `free` is an id like any other in the app's list
  (`patterns`), in a person's list (`Brief.patterns`, the `layouts` argument of the chat tools) and in the list a
  deck keeps (`DeckSpec.patterns`).

The line tells the model to use a free slide only for content that no layout shows, and how many parts read
well at the template's text size. While free slides are on offer no layout is made for the deck
(`composeLayouts`).

**Which kinds.** What the deck has decides what a slide may be planned with:

| Kind | Offered when |
| --- | --- |
| `text`, `list`, `card`, `step`, `quote`, `note` | Always. In a text-only deck: `text`, `list`, `quote` and `note` alone. |
| `figure`, `chart`, `table` | The brief holds digits, or `numberPatterns` is `always`. |
| `photo` | `imageDensity` is not `none`. A plan gives one slide at most six photos to generate, and one more for each photo supplied in `assets`. Of the photos of a slide the writer may describe six; the others take a supplied photo. |
| `graphic` | `assets` has a picture of kind `graphic`. |

A card is offered an icon only when `icons` is given.

**In a plan.** A slide of a plan that is a free slide lists its parts, and how they are set out where the plan
says so. These two fields are all a free slide adds to a plan:

```json
{ "pattern": "free", "purpose": "Every store at a glance", "blocks": ["chart", "chart", "chart", "chart", "note"], "arrange": { "columns": 2 } }
```

A model names the parts in `shows`, as runs of like parts with the `count` of each ("4 charts, then 1 note"):

```json
{ "pattern": "free", "purpose": "Every store at a glance", "shows": [{ "count": 4, "kind": "chart" }, { "count": 1, "kind": "note" }], "shows_main": "auto", "side_by_side": 2 }
```

An answer that gives one entry for each part and numbers them (`count` 1, 2, 3, 4 for four charts) is read as
four charts, not ten. A slide is planned with at most 400 parts, and `writeDeck` refuses a plan that gives a
slide more, with a `PlanError` that names the slide, before any model is asked. `editPlan` shows the slide with
its parts and keeps them unless the request changes what the slide shows. From 0.6.3 a slide whose text from a
document would be cut beside the parts the plan had, or whose parts leave it under half of the words asked, is
its text alone, set dense; a part a change asks for stays, in that change and the ones after it.

**Written.** `writeDeck` works the slide's slots out for exactly these parts, each with the room it has, and
records the frame the slide was planned on in `base`. Up to nine parts are written with the rest of the deck.
A slide of more parts is written by calls of its own, cut into ranges that each fit one answer, three calls at
a time: about a minute of model time for every 50 charts. A call that writes a later range is told which of
their kind its parts are (`the 31st to the 60th chart of the slide`). Every number a chart shows must appear in
the brief; the check does not tell whether it stands in the right chart.

What the writer reads of a slot is the slot's own description and then its limits, each once (see "What a slot
says before the slide is written"): the planner changes neither. An array of like parts says what the smallest
of them holds. Text more than a quarter over the budget it was shown goes back for repair, for every budget of
12 characters or more.

**Parts without content.** A chart, a table or a figure the brief has no numbers for is left empty by the
writer and taken off the slide. So is a part whose content is still missing, or cannot be shown, after the last
repair. The parts that stay are numbered anew and the deck is told:

| Message | Severity | Meaning |
| --- | --- | --- |
| `slide 3: 13 parts had no content and were left out` | warning, on the deck | Nothing was written for these parts. A table with headings and no rows counts as nothing. |
| `slide 3: 1 part could not be shown and was left out` | warning, on the deck | What was written for the part was no chart or table that can be drawn. |
| `slide 3: 1 photo was left out, since at most 6 are made for one slide` | warning, on the deck | The writer described more photos than are generated for one slide, and still did after being asked to use supplied photos for the others. |

A slide left with no part says its purpose as one text; where it has none, the title that was written for it,
or else the title of the deck. A free slide does not stop a deck from being made.

**When a deck is changed.** `planEdit` and `editDeck` change a free slide part by part. The model that plans the
change is shown the slide's parts by number, without their data, and names the parts the slide has afterwards:
which existing part each one is, or that it is new, and whether its content is to be written again.

| The change | What happens |
| --- | --- |
| It leaves the slide alone | The slide is copied as it is. |
| It sets the slide out differently (two side by side, the first part on top) | Code sets `arrange`. Nothing is written. |
| It takes parts off, or puts them in another order | Code makes the slide from the parts that stay, under their new numbers. Nothing is written, unless the plan has words for the slide: then one small call writes its title, as for "take the note off and call the slide Sales 2025". |
| It adds parts, or changes the content of some | Those parts are written. Every other part is carried over by code, and no model is shown its content. |
| It changes the content of the slide and not its parts (a translation) | Every part is written again from what it holds. Pictures stay unless the change is about them. |
| It has words for the slide and names no part | Every part is offered to the writer with what it holds. The writer may return the parts that change and no others: a part that has content and was not named to be written again keeps what it had, and no repair round is spent on the number of items. |
| It gives a slide of a pattern more than the pattern shows | The slide becomes a free slide. Its chart and its table go to the first new parts of those kinds that the plan does not ask to be written again, and its photos to the first photo parts unless the change is about pictures. A chart the plan asks to be written again is written by the model from the slide's content. |

**What a part that is written again is offered.** The slide is planned again for the parts it has after the
change, and with the optional parts it holds: a caption under its one photo or a heading over its one text
is counted in, although the rules plan neither for such a slide. It keeps its room, and is offered again when
its part is written again. A photo that is replaced is offered its caption with the new picture.

A chart or a table the slide holds may be larger than its place reads well at (three series where a new chart
of that size is told two). It is offered as it is, with limits no lower than what it holds, and is not sent
back for its size when it comes back no larger. The check of the deck says the same of it before the change
and after: `the chart has 3 series; the slot fits 2`. Like charts of a slide share one value axis, so new
numbers for one chart can change the scale its neighbours are drawn on; their numbers stay as they are.

**How the slide is set out after a change.** The plan of a change says for every slide where the main part
stands (`auto`, `left`, `right`, `top`, `none`) and how many parts stand side by side (0 for the software's
choice), and is told to repeat the current values unless the request is about the arrangement.

| The plan gives | On a slide it changes (it has words for it, or its parts change) | On a slide it leaves alone |
| --- | --- | --- |
| `left`, `right`, `top`, `none`, or a number from 1 | Set. | Set, by code, with nothing written. |
| `auto`, or 0 | The slide is handed back to the software's own choice: that part of `arrange` is removed. | Nothing: the slide is copied as it is. |
| Nothing (the property is left out) | The slide stays set out as it was. | The slide is copied as it is. |

`auto` and 0 are also what a model answers for a slide it has nothing to say of, so on such a slide they are
not taken as a wish. To hand a slide back to the software, the request has to change something on it too.

**The title in a change.** A slide's title is offered to the writer with the parts that are written, and the
writer is told to repeat the current title unless the request is about it. A title the writer leaves out stays.
On a slide of up to nine parts an empty title (`""`) takes the title off; on a larger slide it keeps the one the
slide had (see "Limits").

A plan that has words for a slide while it names every part as it is, none to be written again, can only mean
the title, which is then offered alone. When that title comes back as it was, nothing on the slide changed
although the request spoke of it, and the deck is told in a warning without a slide number:
`slide 3: nothing on it was changed: the plan of the change named no part to write again`. No call is added.

A part that is to be written again and comes back empty, or with content that cannot be shown, stays as it was:
a change takes no part off a slide by leaving it unsaid. What the plan asked for outright is asked for again
first, like any missing content. The deck is then told in a warning without a slide number:
`slide 3: nothing was written for 1 part, and it stays as it was`, or `slide 3: what was written for 1 part
could not be shown, and it stays as it was`. A new part that comes back without content is left out, with the
warning above. See "Free slides in an edit" in [generating.md](generating.md).

**In the chat tools.** A plan shows such a slide with `"layout": "free"`, `"free_layout": true` and a
description of what it holds; a presentation adds `"parts": "6 charts and a note"`. See
[chat-tools.md](chat-tools.md).

A plan or a deck that has a free slide is written, changed and rendered whether or not free slides are on offer:
the conditions above decide only what a new slide may be.

## Limits

- A free slide is laid out by rules, not by taste. It is tidy and in the template's look; it is not the work
  of a designer, and a designed pattern is the better slide wherever one fits.
- Blocks are set out in the order given. There is no free placement by coordinates.
- A free slide has no row of items to vary: its blocks are its items. `items` is ignored on it.
- The title of a free slide of ten parts or more is changed, not removed, by an edit (`planEdit`, `editDeck`,
  the chat tools): such a slide is written by calls of its own, in which an empty title keeps the one the slide
  had. Remove it in the deck JSON, or with `freeContent` and `freeSlide`.
- A template whose patterns have no running text that reads on the frame's background cannot carry free
  slides; `validateDeck` says so.
- The styles of a free slide are copied from text boxes that state their own look. Text in a placeholder takes
  its look from the slide layout and cannot be copied: a template whose running text is all in placeholders
  cannot carry free slides, and `validateDeck` says that too. A title in a placeholder is fine.
- A line of text is counted from its size and the line spacing of the shape it is a copy of. A line spacing
  under 100% is counted as 100%, so such text leaves a little of its room unused. The space a shape sets
  before and after its paragraphs is not copied: the paragraphs of a free slide are half their size apart, in
  whole points, or closer where spacing gives way (see "How large things are"). Where lines close up, the copy
  states its line spacing in whole percents, or in whole points where its source states points, as PowerPoint
  draws them.
- A big number (a figure's value, the number of a step) is counted as one line of its own size, whatever line
  spacing its shape has.
- Text is sized by character counts unless the template is loaded with fonts (see [text-fit.md](text-fit.md)).
- A free slide is not shown in [preview.md](preview.md)'s pictures of a template's patterns: it has no fixed
  look to show before it has content. A deck's free slides are drawn like its other slides.
