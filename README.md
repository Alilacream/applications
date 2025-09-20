Application Repository

Welcome to my application repository 👋.
Here, I challenge myself to build projects that either help people I know or tackle ideas that truly interest me.

Project Overview

This project is inspired by a big data Excel management file for engineers.
The original file contains many columns, but the most important ones are:

N° PRIX

DESIGNATION

UNITE

QUANTITES

P.U DH.HT

TOTAL

DESCRIPTION

The issue: the DESCRIPTION column takes up a lot of space, making the table difficult to read.

The Solution

This application processes the big Excel file and splits it into two outputs:

Excel File → Contains all the main calculation columns (N° PRIX, DESIGNATION, UNITE, QUANTITES, P.U DH.HT, TOTAL) but without the description, making it much more compact and readable.

Word File → Contains the pair (N° PRIX, DESIGNATION) followed by the DESCRIPTION, so all the verbose details are stored neatly in a separate document.

This way, engineers can work easily with the Excel file for calculations while still keeping the descriptions accessible when needed.

Personal Note

I originally coded this application for my brother ❤️.
I call it “vibe coding” because I started it quickly, but unlike usual vibe coding, I actually took my time analyzing and improving:

Exporting/importing functions

Error handling

HTML tag threading

This is Version 1 of the app.
Next steps: I plan to refactor and split the big code chunk into multiple clean components.

Wish me luck on my journey 🚀, and I wish the same for every developer reading this repo!

Do you want me to also add a “How to Use” section (with installation steps, commands, dependencies) so that others could actually run your app?
