# Camel Type: Python Coding Guidelines

Camel Type is a coding style optimized for screen reader productivity and systematic readability. The following rules apply to Python 3.

---

## 1. Variable and Argument Naming

Use Hungarian prefix notation to indicate type. Prefix rules:

- `a` — array (or tuple used as a fixed sequence)
- `b` — boolean
- `bin` — binary buffer (`bytes` or `bytearray`)
- `dt` — date-time (`datetime` object)
- `f` — file object
- `i` — integer
- `l` — list
- `n` — real number (float)
- `s` — string
- `d` — dictionary
- `v` — variant (unknown or mixed type)

For **Python object instances** (class instances, library objects), use the **lowercase class name** as the prefix. If there is only one instance of that class in scope, the class name prefix is the entire variable name — no additional suffix is added.

Examples:
- One `Path` instance: `path`
- One `requests.Session` instance: `session`
- One `BeautifulSoup` instance: `beautifulSoup`
- A caught exception: `error`

---

## 2. Constant Naming

Constants use the **same lower camel case naming convention as variables**, including the Hungarian type prefix. They are distinguished only by intent and placement (defined at the top of the module), not by any difference in capitalization or formatting.

Examples: `sDefaultEncoding = "utf-8"`, `iTimeoutSec = 60`

---

## 3. Capitalization

Use **lower camel case** for all custom names: variables, constants, function names, and parameters. Python conventionally uses snake case, but Camel Type overrides this with lower camel case everywhere except where a framework or external API requires otherwise.

Class names, if defined, use upper camel case as Python convention requires.

---

## 4. Variable Initialization

Python does not require explicit variable declarations. Assign zero-value initializations at the **top of each function scope**, grouped by type in alphabetical order of prefix letter, variables of the same type assigned alphabetically.

Where a meaningful initial value is not yet known, assign a sensible zero-value: `0` for integers, `0.0` for floats, `""` for strings, `[]` for lists, `{}` for dicts, `None` for objects.

Example:
```python
iCount = 0
iMax = 100
sDir = ""
sTitle = ""
sUrl = ""
```

---

## 5. Constants

Define constants at the **top of the module**, before any function definitions, grouped by type, each group in alphabetical order, constants within each group in alphabetical order.

Example:
```python
iTimeoutSec = 60
sEncoding = "utf-8"
sFallbackTitle = "untitled-page"
sResultsDir = "results"
```

---

## 6. Script Structure and Entry Point

Prefer **top-level script code with no indentation** rather than wrapping everything in a `main` function. Write the main logic directly at the module's top level so a screen reader user can navigate to the bottom of the file to find the most recently written code.

When a `main` function is used (e.g., to enable the `if __name__` guard for reusable modules), place it **last** in the file, after all other functions. The `if __name__ == "__main__":` block appears at the very bottom, calling `main()` on a single line.

```python
if __name__ == "__main__": main()
```

---

## 7. Function Order

All functions other than `main` are listed in **strict alphabetical order**. If a `main` function is used, it appears **last**. There are no organizational sections or divider comments grouping functions — strict alphabetical order is sufficient.

---

## 8. Functions

- Define all routines as **functions**. Every function should return a value, even if `None`.
- Use simple **single-line conditionals** for one-consequence cases:
  ```python
  if not sUrl: raise ValueError(sUsage)
  if bReady: return True
  ```

---

## 9. Loops

Prefer **for-each** style loops over index-based loops whenever iterating a collection.

```python
for sItem in lItems: process(sItem)
```

Use `enumerate()` only when the index itself is needed. Avoid `range(len(...))`.

---

## 10. Imports

Use exactly **two import statements**: one for built-in standard library modules, one for third-party modules. Each statement lists module names in **alphabetical order** on a single line, separated by commas. Place the built-in import first, then the third-party import, separated by a blank line.

```python
import json, os, pathlib, re, sys

import bs4, playwright, requests
```

If `from X import Y` syntax is needed for specific names, place those lines after the two main import statements, alphabetized.

---

## 11. String Delimiters

Use **double quotes** for all strings. Use f-strings when embedding expressions.

```python
sPath = f"{sResultsDir}/{sTitle}"
```

---

## 12. Magic Numbers

Never use literal numeric or string values in logic. Assign them to named constants and reference the constant name.

---

## 13. Object and Array Literals

Keep short dict and list literals on a **single line** where practical, to maximize information density per line for screen reader review.

```python
dConfig = {"encoding": sEncoding, "timeout": iTimeoutSec}
lEngines = ["axe", "ibm", "alfa"]
```

---

## 14. Error Handling

Use `try`/`except` with a named exception variable. Catch specific exception types where possible rather than a bare `except`.

```python
try:
    result = riskyOperation()
except ValueError as error:
    print(f"Validation error: {error}")
    sys.exit(1)
```
