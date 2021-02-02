import React, { useState, useRef, useEffect } from "react"
import PropTypes from "prop-types"
import "./styles.css"
import Util from "./util"

const DEBUG = false
const DISPLAY_SHEET_JSON = true

// Keep in mind that this feature might be generally faster than a full
// cloning, but tends to make garbage collection harder.
const ENABLE_SELECTIVE_CLONING = false

// https://en.wikipedia.org/wiki/Observer_pattern
class Cell {
  constructor(refOrCoords, value, sheet) {
    this.ref = Util.asRef(refOrCoords)
    this.sheet = sheet
    this.subjects = new Set() // Cells who I watch for changes.
    this.observers = new Set() // Cells who watch me for changes.
    this.modified = false // Changes when value is changed or (re)evaluated.
    this.evaluated = false // Changes when value is (re)evaluated.
    this.invalid = false

    this.setValue(value)
  }

  static create() {
    return Object.create(this.prototype)
  }

  clone(attrsToMerge = {}) {
    const clone = Cell.create()

    // Copy all object properties.
    Object.keys(this).forEach((property) => {
      clone[property] =
        this[property] instanceof Set ? new Set(this[property]) : this[property]
    })

    Object.entries(attrsToMerge).forEach(([key, value]) => {
      clone[key] = value
    })

    return clone
  }

  hasFormula() {
    return Util.isFormula(this.value)
  }

  setValue(newValue, descendantObservers) {
    if (this.value !== newValue) {
      this.value = newValue
      this.modified = true

      // Assume cell is valid.
      this.invalid = false
      this.errorMessage = null

      const oldSubjects = new Set(this.subjects)
      const newSubjects = Util.isFormula(newValue)
        ? this._extractCells(Util.expandRanges(newValue))
        : new Set()

      if (this._hasCircularDependency(newSubjects)) {
        this.evaluatedValue = null
        this.invalid = true
        this.errorMessage = "Circular dependency"
      }

      if (!this.invalid) {
        this.subjects = newSubjects

        const subjectsAdded = Util.setDiff(this.subjects, oldSubjects)
        const subjectsRemoved = Util.setDiff(oldSubjects, this.subjects)

        subjectsAdded.forEach((ref) => {
          this.sheet.findOrCreateCell(ref)._registerObserver(this.ref)
        })

        subjectsRemoved.forEach((ref) => {
          this.sheet.findOrCreateCell(ref)._unregisterObserver(this.ref)
        })

        descendantObservers ??= this.descendantObservers()

        if (DEBUG)
          console.log(`${this.ref}'s descendantObservers: `, [
            ...descendantObservers
          ])

        this._evaluateValue(descendantObservers)
      }
    }
  }

  descendantObservers(visited = []) {
    // Please notice the cell itself (this) will be included in the resulting collection.
    return Util.setAppend(
      new Set([this.ref]),
      [...this.observers].flatMap((ref) => {
        if (visited.indexOf(ref) >= 0) return [] // Cell already processed.

        visited.push(ref)

        return [...this.sheet.findCell(ref).descendantObservers(visited)]
      })
    )
  }

  copyTo(targetRef) {
    let copiedValue = this.value

    if (Util.isFormula(this.value)) {
      const subjects = this._extractCells(this.value)

      const subjectNewRefs = [...subjects].reduce((acc, ref) => {
        acc[ref] = Util.newRefForCopy(ref, this.ref, targetRef)

        return acc
      }, {})

      copiedValue = Util.templateReplaceForCopy(this.value, subjectNewRefs)
    }

    // Assure all cells are properly re-evaluated (when applicable) by
    // setting their `evaluated` flag to false prior to reprocessing them.
    Object.values(this.sheet.cells).forEach((cell) => (cell.evaluated = false))

    const newCell = this.sheet.updateOrCreateCell(targetRef, copiedValue)

    return newCell
  }

  moveTo(targetRef, sourceRefs) {
    if (DEBUG) console.log(`Moving ${this.ref} to ${targetRef}`)

    this.sheet.findOrCreateCell(targetRef).observers.forEach((ref) => {
      const observer = this.sheet.findCell(ref)

      observer.evaluated = false
    })

    // Copy source cell raw/original value to current ref.
    const targetCell = this.sheet.updateOrCreateCell(targetRef, this.value)

    // For each observer of the source cell, update its references
    // to the subject's new position.
    this.observers.forEach((ref) => {
      const observer = this.sheet.findCell(ref)
      let observerValue = observer.value
      const ranges = new Set(Util.textScan(observerValue, Util.RANGE_REGEXP))

      ranges.forEach((range) => {
        const rangeCells = new Set(Util.expandRange(range).flat(2))
        const rangeCellsNotInMove = Util.setDiff(
          rangeCells,
          new Set(sourceRefs)
        )

        // Are all cells within the range being moved in the same operation?
        if (rangeCellsNotInMove.size === 0) {
          if (DEBUG)
            console.log(`Observer ${ref}: entire range ${range} being moved`)

          const [fromRef, toRef] = range.split(":")
          const newFromRef = Util.newRefForCopy(fromRef, this.ref, targetRef)
          const newToRef = Util.newRefForCopy(toRef, this.ref, targetRef)

          observerValue = observerValue.replace(
            new RegExp(`\\b${range}\\b`, "g"),
            [newFromRef, newToRef].join(":")
          )
        }
      })

      if (DEBUG)
        console.log(
          `Updating ${ref} references with current value ${observerValue}`
        )

      observer.setValue(
        Util.templateReplaceForMove(observerValue, {
          [this.ref]: targetRef
        })
      )

      observer.evaluated = false
    })

    // Assure the cell being moved is properly re-evaluated by setting
    // its `evaluated` to false prior to nullify it.
    this.evaluated = false
    this.setValue(null)

    return targetCell
  }

  ///////////////////////
  // Private functions //
  ///////////////////////

  _evaluateValue(updatedCellDescendantObservers) {
    if (DEBUG) console.log(`Evaluating ${this.ref}...`)

    if (this.evaluated) {
      if (DEBUG) console.log(`Cell ${this.ref} already evaluated. Skipping...`)

      return
    }

    this.evaluated = true

    // Assume cell is valid.
    this.invalid = false
    this.errorMessage = null

    const previousValue = this.evaluatedValue
    let parsedFormula

    if (Util.isFormula(this.value)) {
      try {
        parsedFormula = this._parseFormula()
      } catch (e) {
        if (e instanceof Util.InvalidRefInFormula) {
          this.evaluatedValue = null
          this.invalid = true
          this.errorMessage = e.message
        } else {
          throw e
        }
      }

      if (!this.invalid) {
        try {
          /* eslint-disable no-unused-vars */
          const SUM = (...values) =>
            values.flat(2).reduce((acc, i) => acc + i, 0)
          const COUNT = (...values) => values.flat(2).length
          const AVG = (...values) => {
            const flattenedValues = values.flat(2)
            return SUM(...flattenedValues) / COUNT(...flattenedValues)
          }
          const MAX = (...values) => Math.max(...values.flat(2))
          const MIN = (...values) => Math.min(...values.flat(2))
          const ROWS = (values) => values.length
          const COLS = (values) => (values[0] ?? []).length
          /* eslint-disable no-unused-vars */

          this.evaluatedValue = eval(parsedFormula) // eslint-disable-line no-eval
        } catch {
          this.evaluatedValue = null
          this.invalid = true
          this.errorMessage = "Invalid formula"
        }
      }
    } else {
      this.evaluatedValue = this.value
    }

    if (DEBUG)
      if (this.evaluatedValue !== previousValue) {
        console.log(`${this.ref} new value: ${this.evaluatedValue}`)
      } else {
        console.log(`No changes to ${this.ref}'s value!`)
      }

    if (this.evaluatedValue !== previousValue) {
      this.modified = true

      this._notifyObservers(updatedCellDescendantObservers)
    }
  }

  _parseFormula(defaultValue = 0) {
    const subjectValues = [...this.subjects].reduce((acc, ref) => {
      const cell = this.sheet.findCell(ref)

      if (cell.invalid)
        throw new Util.InvalidRefInFormula(`Invalid ref [${ref}]`)

      acc[ref] = cell._valueForFormulaCalculation(defaultValue)

      return acc
    }, {})

    const parsedFormula = Util.rawFormula(
      // Replace all (subject) references in formula.
      Util.templateReplace(
        Util.expandRanges(this.value), // e.g. '=A1+B2'
        subjectValues // e.g. { A1: 10, B2: 20 }
      )
    )

    if (DEBUG)
      console.log(
        `${this.ref} -> formula: '${Util.rawFormula(
          this.value
        )}', subjects: ${JSON.stringify(
          subjectValues
        )}, replacedFormula: '${parsedFormula}'`
      )

    return parsedFormula
  }

  _valueForFormulaCalculation(defaultValue) {
    if (Util.isEmpty(this.value)) return defaultValue
    else if (this.hasFormula()) return this.evaluatedValue
    else if (Object.is(Number(this.value), NaN)) return defaultValue
    else return this.evaluatedValue
  }

  _registerObserver(observer) {
    this.observers.add(observer)
  }

  _unregisterObserver(observer) {
    this.observers.delete(observer)
  }

  _updateObserver(updatedCellDescendantObservers) {
    this._evaluateValue(updatedCellDescendantObservers)
  }

  _notifyObservers(updatedCellDescendantObservers) {
    if (DEBUG)
      console.log(
        `Notifying ${this.ref}'s observers: [${[...this.observers].join(", ")}]`
      )

    this.observers.forEach((ref) => {
      const observer = this.sheet.findCell(ref)
      const commonCells = Util.setIntersect(
        observer.subjects,
        updatedCellDescendantObservers
      )
      const nonEvaluatedCommonCells = [...commonCells].filter(
        (commonCellRef) => !this.sheet.findCell(commonCellRef).evaluated
      )

      if (nonEvaluatedCommonCells.length > 0) {
        if (DEBUG)
          console.log(
            `Skipping observer ${ref} due to non-evaluated common cell(s) [${nonEvaluatedCommonCells.join(
              ", "
            )}]`
          )

        return
      }

      observer._updateObserver(updatedCellDescendantObservers)
    })
  }

  _extractCells(value) {
    return Util.findRefsInFormula(value)
  }

  _hasCircularDependency(subjects, visited = []) {
    return (
      subjects.has(this.ref) ||
      [...subjects].some((ref) => {
        if (visited.indexOf(ref) >= 0) return false // Cell already processed.

        visited.push(ref)

        const subject = this.sheet.findOrCreateCell(ref)

        return this._hasCircularDependency(subject.subjects, visited)
      })
    )
  }
}

class Sheet {
  constructor(initialCellData) {
    this.id = Sheet.generateId()
    this.cells = {}

    this._loadInitialCellData(initialCellData)
  }

  static create() {
    return Object.create(this.prototype)
  }

  static generateId() {
    return Date.now()
  }

  clone(cellsToClone) {
    const clone = Sheet.create()

    clone.id = Sheet.generateId()

    clone.cells = Object.entries(this.cells).reduce((acc, [ref, cell]) => {
      acc[ref] =
        !ENABLE_SELECTIVE_CLONING || cellsToClone.has(ref)
          ? cell.clone({ modified: false, evaluated: false, sheet: clone })
          : cell // Notice that this (reused) cell will still point to its old sheet.

      return acc
    }, {})

    return clone
  }

  dimensions() {
    let [maxRow, maxCol] = [0, 0]

    Object.keys(this.cells).forEach((ref) => {
      const { row, col } = Util.asCoords(ref)
      if (row > maxRow) maxRow = row
      if (col > maxCol) maxCol = col
    })

    return [maxRow, maxCol]
  }

  cellCount() {
    return Object.keys(this.cells).length
  }

  findCell(refOrCoords) {
    const ref = Util.asRef(refOrCoords)

    return this.cells[ref]
  }

  findOrCreateCell(refOrCoords, valueIfNew) {
    const ref = Util.asRef(refOrCoords)

    const cell = (this.cells[ref] =
      this.cells[ref] ?? new Cell(ref, valueIfNew, this))

    return cell
  }

  // upsert operation
  updateOrCreateCell(refOrCoords, valueIfNew, descendantObservers) {
    let cell = this.findCell(refOrCoords)

    if (cell) {
      cell.setValue(valueIfNew, descendantObservers)
    } else {
      const ref = Util.asRef(refOrCoords)

      cell = this.cells[ref] = new Cell(ref, valueIfNew, this)
    }

    return cell
  }

  ///////////////////////
  // Private functions //
  ///////////////////////

  _loadInitialCellData(initialCellData) {
    Object.entries(initialCellData).forEach(([ref, value]) => {
      this.updateOrCreateCell(ref, value)

      // Reset all already created cell's `evaluated` flag, so the
      // optimizatin check inside `Cell#_evaluateValue` works as expected.
      Object.values(this.cells).forEach((cell) => {
        cell.evaluated = false
      })
    })
  }
}

// // Euler's number calculation [https://en.wikipedia.org/wiki/E_(mathematical_constant)].
// const initialCellData = Util.sequenceReduce(
//   18 - 1,
//   (acc, i) => ({
//     ...acc,
//     [`A${i + 4}`]: i + 1,
//     [`B${i + 4}`]: `=B${i + 3}*${i + 1}`,
//     [`C${i + 4}`]: `=1/B${i + 4}`,
//     [`D${i + 4}`]: `=D${i + 3}+C${i + 4}`
//   }),
//   {
//     A1: "N",
//     B1: "N!",
//     C1: "1 / N!",
//     D1: "E = Σ (1 / N!), 0 <= N < ∞",
//     A3: 0,
//     B3: 1,
//     C3: "=1/B3",
//     D3: "=C3"
//   }
// )

// // Standard factorial sequence.
// const initialCellData = Util.sequenceReduce(
//   10 - 1,
//   (acc, i) => ({
//     ...acc,
//     [`A${i + 2}`]: `=A${i + 1} * ${i + 2}`
//   }),
//   {
//     A1: 1
//   }
// )

// initialCellData = {
//   A1: 1,
//   B1: "=A1+C2+D4",
//   C1: "=B1+A2",
//   A2: "=A1+B1",
//   B2: "=A1+D4",
//   C2: "=B2+D4",
//   A3: "=A2+C1+C3",
//   B3: "=B2+1",
//   C3: "=B2+1",
//   C4: "=A1+1",
//   D4: "=A1+1"
// }

// const initialCellData = {
//   A1: 1,
//   B1: 2,
//   A2: 3,
//   B2: 4,
//   C1: "=SUM(A1:B1)",
//   C2: "=SUM(A2:B2)",
//   A3: "=SUM(A1:A2)",
//   B3: "=SUM(B1:B2)",
//   C3: "=SUM(A1:B2)"
// }

// const initialCellData = Util.generateCellSquares(4, "A1", 1)

const initialCellData = Util.generateSpiralSequence(
  10,
  "south",
  "left",
  [{ A1: 1 }],
  (_i, previousRefs, _nextRef) => `=${previousRefs[previousRefs.length - 1]}+1`
)

// const initialCellData = Util.generateSpiralSequence(
//   5,
//   "south",
//   "left",
//   [{ A1: 1 }],
//   (_i, previousRefs, _nextRef) => `=${previousRefs[previousRefs.length - 1]}+1`
// )

// // Spiral reversed sequence.
// const initialCellData = Util.generateSpiralSequence(
//   6,
//   "east",
//   "right",
//   [{ A1: "=B1+1" }],
//   (_i, _previousRefs, nextRef) => `=${nextRef}+1`
// )

// // Spiral factorial.
// const initialCellData = Util.generateSpiralSequence(
//   5,
//   "east",
//   "right",
//   [{ A1: 1 }],
//   (i, previousRefs, _nextRef) => `=${previousRefs[previousRefs.length - 1]}*${i + 2}`
// )

// // Spiral Fibonacci sequence.
// const initialCellData = Util.generateSpiralSequence(
//   10,
//   "south",
//   "left",
//   [{ A1: 0 }, { A2: 1 }],
//   (_i, previousRefs, _nextRef) =>
//     `=${previousRefs[previousRefs.length - 2]}+${
//       previousRefs[previousRefs.length - 1]
//     }`
// )

// // Spiral "generalized" Fibonacci sequence (adding last 3 items, instead of 2).
// const initialCellData = Util.generateSpiralSequence(
//   10,
//   "east",
//   "right",
//   [{ A1: 0 }, { B1: 0 }, { C1: 1 }],
//   (_i, previousRefs, _nextRef) => `=${previousRefs[previousRefs.length - 3]}+${previousRefs[previousRefs.length - 2]}+${previousRefs[previousRefs.length - 1]}`
// )

const initialSheet = new Sheet(initialCellData)

const Spreadsheet = (props) => {
  const MAXIMUM_CELLS = 10000
  const HEADER_LIMITS = { rows: 15, cols: 5 }

  const [sheet, setSheet] = useState(initialSheet)
  const [clipboard, setClipboard] = useState({ range: null, action: null })
  const [selectedRangeCorner1, setSelectedRangeCorner1] = useState(null)
  const [selectedRangeCorner2, setSelectedRangeCorner2] = useState(null)
  const cellsRef = useRef({})

  const dimensions = sheet.dimensions()

  if (dimensions[0] * dimensions[1] > MAXIMUM_CELLS) {
    return (
      <h3>
        Sorry, only sheets with a maximum of {MAXIMUM_CELLS} cells are
        supported.
      </h3>
    )
  }

  // useEffect(() => {
  //   _gotoCell("A1")
  // }, [])

  let selectedRange

  if (selectedRangeCorner1 && selectedRangeCorner2) {
    const { row: corner1Row, col: corner1Col } = Util.asCoords(
      selectedRangeCorner1
    )
    const { row: corner2Row, col: corner2Col } = Util.asCoords(
      selectedRangeCorner2
    )

    let rangeCorners

    if (corner1Row <= corner2Row) {
      if (corner1Col <= corner2Col) {
        // Corner 2 is below (or at the same row) and to the right (or at the same col) of corner 1.
        rangeCorners = [selectedRangeCorner1, selectedRangeCorner2]
      } else {
        // Corner 2 is below (or at the same row) and to the left of corner 1.
        rangeCorners = [
          Util.asRef([corner1Row, corner2Col]),
          Util.asRef([corner2Row, corner1Col])
        ]
      }
    } else {
      if (corner1Col <= corner2Col) {
        // Corner 2 is above and to the right (or at the same col) of corner 1.
        rangeCorners = [
          Util.asRef([corner2Row, corner1Col]),
          Util.asRef([corner1Row, corner2Col])
        ]
      } else {
        // Corner 2 is above and to the left of corner 1.
        rangeCorners = [selectedRangeCorner2, selectedRangeCorner1]
      }
    }

    selectedRange = rangeCorners.join(":")
  }

  const selectedRangeRefs = selectedRange
    ? Util.expandRange(selectedRange).flat(2)
    : []

  const handleCellInputKeyDown = (ref) => (event) => {
    const { target, key } = event
    const isTextSelected = target.selectionStart < target.selectionEnd
    const isTextFullySelected =
      target.selectionStart === 0 && target.selectionEnd === target.value.length
    const input = _cellInput(ref)
    const cell = sheet.findCell(ref)
    const { row, col } = Util.asCoords(ref)

    const firstCell = [1, 1]
    const lastCell = dimensions
    const firstRow = 1
    const lastRow = dimensions[0]
    const firstCol = 1
    const lastCol = dimensions[1]

    const isFirstCell = row === firstCell[0] && col === firstCell[1]
    const isLastCell = row === lastCell[0] && col === lastCell[1]
    const isFirstRow = row === firstRow
    const isLastRow = row === lastRow
    const isFirstCol = col === firstCol
    const isLastCol = col === lastCol

    let destinationCoords

    switch (key) {
      case "c": // (Command|Ctrl)+c (COPY operation)?
      case "x": // (Command|Ctrl)+x (CUT operation)?
        // Proceed only if Command or Control pressed.
        if (!event.metaKey && !event.ctrlKey) return

        // Proceed only if text is fully selected.
        if (!isTextFullySelected) return

        event.preventDefault()

        // Copy currently selected range to clipboard.
        setClipboard({
          range: selectedRange,
          action: { c: "copy", x: "cut" }[key]
        })

        break

      case "v": // (Command|Ctrl)+v pressed (PASTE operation)?
        // Proceed only if Command or Control pressed.
        if (!event.metaKey && !event.ctrlKey) return

        // Proceed only if clipboard not empty.
        if (!clipboard.range || !clipboard.action) return

        event.preventDefault()

        const sourceRefs = Util.expandRange(clipboard.range)

        setSheet((previousSheet) => {
          const sheetClone = previousSheet.clone()

          sourceRefs.forEach((sourceRefsRow, rowIndex) => {
            sourceRefsRow.forEach((sourceRef, colIndex) => {
              const sourceCell = sheetClone.findCell(sourceRef)

              if (!sourceCell) return

              const targetRef = Util.asRef([row + rowIndex, col + colIndex])
              let targetCell

              switch (clipboard.action) {
                case "copy":
                  targetCell = sourceCell.copyTo(targetRef)
                  break

                case "cut":
                  targetCell = sourceCell.moveTo(targetRef, sourceRefs.flat(2))

                  _syncCellInput(sheetClone, sourceRef, sourceCell)

                  // Sync all observers' inputs.
                  targetCell.observers.forEach((observerRef) => {
                    _syncCellInput(sheetClone, observerRef)
                  })
                  break

                default:
              }

              _syncCellInput(sheetClone, targetRef, targetCell)
            })
          })

          return sheetClone
        })

        if (clipboard.action === "cut") {
          // Clear the clipboard.
          setClipboard({ range: null, action: null })
        }

        break

      case "Escape":
        if (isTextFullySelected) {
          // Allow the user to start editing the cell during navigation.
          document.getSelection().collapseToEnd()
        } else {
          // Restore value.
          input.value = cell?.value ?? ""
        }
        break

      case "Enter":
        event.preventDefault()

        input.blur()

        if (isLastCell) {
          destinationCoords = firstCell
        } else if (isLastRow) {
          destinationCoords = [firstCol, col + 1]
        } else {
          destinationCoords = [row + 1, col]
        }

        const destinationRef = Util.asRef(destinationCoords)

        _gotoCell(destinationRef)

        break

      case "ArrowDown":
      case "ArrowUp":
      case "ArrowLeft":
      case "ArrowRight":
        if (event.shiftKey) {
          // Proceed only if text is fully selected.
          if (!isTextFullySelected) return

          event.preventDefault()

          let { row: corner2Row, col: corner2Col } = Util.asCoords(
            selectedRangeCorner2
          )

          switch (key) {
            case "ArrowDown":
              corner2Row = Math.min(corner2Row + 1, dimensions[0])
              break
            case "ArrowUp":
              corner2Row = Math.max(corner2Row - 1, 1)
              break
            case "ArrowLeft":
              corner2Col = Math.max(corner2Col - 1, 1)
              break
            case "ArrowRight":
              corner2Col = Math.min(corner2Col + 1, dimensions[1])
              break
            default:
          }

          setSelectedRangeCorner2(Util.asRef([corner2Row, corner2Col]))
        } else {
          if (!isTextFullySelected) return

          event.preventDefault()

          input.blur()

          switch (key) {
            case "ArrowDown":
              if (isLastCell) {
                destinationCoords = firstCell
              } else if (isLastRow) {
                destinationCoords = [firstRow, col + 1]
              } else {
                destinationCoords = [row + 1, col]
              }

              break

            case "ArrowUp":
              if (isFirstCell) {
                destinationCoords = lastCell
              } else if (isFirstRow) {
                destinationCoords = [lastRow, col - 1]
              } else {
                destinationCoords = [row - 1, col]
              }

              break

            case "ArrowLeft":
              if (isFirstCell) {
                destinationCoords = lastCell
              } else if (isFirstCol) {
                destinationCoords = [row - 1, lastCol]
              } else {
                destinationCoords = [row, col - 1]
              }

              break

            case "ArrowRight":
              if (isLastCell) {
                destinationCoords = firstCell
              } else if (isLastCol) {
                destinationCoords = [row + 1, firstCol]
              } else {
                destinationCoords = [row, col + 1]
              }

              break

            default:
          }

          const destinationRef = Util.asRef(destinationCoords)

          _gotoCell(destinationRef)
        }

        break

      case "Tab":
        const isTabInLastCell = () => !event.shiftKey && isLastCell
        const isShiftTabInFirstCell = () => event.shiftKey && isFirstCell

        if (isTabInLastCell() || isShiftTabInFirstCell()) {
          event.preventDefault()

          input.blur()

          const destinationRef = isTabInLastCell()
            ? Util.asRef(firstCell)
            : Util.asRef(lastCell)

          _gotoCell(destinationRef)
        }

        break

      default:
    }
  }

  const handleCellInputBlur = (ref) => (_event) => {
    const cell = sheet.findCell(ref)
    const newValue = Util.normalizeValue(_cellInput(ref).value)

    // Do nothing if cell value hasn't changed.
    if ((!cell && newValue === "") || (cell && newValue === cell.value)) return

    setSheet((previousSheet) =>
      Util.showElapsedTimes(
        () => {
          let descendantObservers
          let targetCells

          if (ENABLE_SELECTIVE_CLONING) {
            descendantObservers = cell?.descendantObservers() ?? new Set()

            targetCells = Util.setAppend(
              descendantObservers,
              Object.entries(previousSheet.cells)
                .filter(([_ref, cell]) => cell.modified || cell.invalid)
                .map(([ref, _cell]) => ref)
            )
          }

          const sheetClone = previousSheet.clone(targetCells)

          // Create new cells on demand.
          sheetClone.updateOrCreateCell(ref, newValue, descendantObservers)

          return sheetClone
        },
        {
          message: `Update cell with ${
            ENABLE_SELECTIVE_CLONING ? "selective" : "full"
          } cloning`,
          // repeat: previousSheet.cellCount() < 1000 ? 10 : 3,
          matchResults: false
        }
      )
    )
  }

  const handleAddNewCol = (event) => {
    _addNewRowOrCol(1, dimensions[1] + 1)
  }

  const handleAddNewRow = (event) => {
    _addNewRowOrCol(dimensions[0] + 1, 1)
  }

  const _addNewRowOrCol = (row, col) => {
    setSheet((previousSheet) => {
      const sheetClone = previousSheet.clone()

      // Create new cells on demand.
      sheetClone.findOrCreateCell([row, col])

      return sheetClone
    })
  }

  const _cellInput = (ref) => cellsRef.current[ref]

  const _gotoCell = (ref, selectText = true) => {
    const cell = _cellInput(ref)

    if (cell) {
      setTimeout(() => {
        cell.focus()
        if (selectText) cell.select()
      }, 0)
    }
  }

  const _syncCellInput = (currentSheet, ref, cell) => {
    if (_cellInput(ref))
      _cellInput(ref).value = (cell ?? currentSheet.findCell(ref)).value
  }

  const handleCellClick = (ref) => (event) => {
    const refInSelectedRange = selectedRangeRefs.indexOf(ref) >= 0

    setSelectedRangeCorner1(ref)
    setSelectedRangeCorner2(ref)

    _gotoCell(ref, !refInSelectedRange)
  }

  const handleSelectAllClick = (_event) => {
    const firstCell = Util.asRef([1, 1])
    const lastCell = Util.asRef(dimensions)

    setSelectedRangeCorner1(firstCell)
    setSelectedRangeCorner2(lastCell)

    _gotoCell(firstCell)
  }

  const handleSelectRowClick = (row) => (_event) => {
    const firstRowCell = Util.asRef([row, 1])
    const lastRowCell = Util.asRef([row, dimensions[1]])

    setSelectedRangeCorner1(firstRowCell)
    setSelectedRangeCorner2(lastRowCell)

    _gotoCell(firstRowCell)
  }

  const handleSelectColClick = (col) => (_event) => {
    const firstColCell = Util.asRef([1, col])
    const lastColCell = Util.asRef([dimensions[0], col])

    setSelectedRangeCorner1(firstColCell)
    setSelectedRangeCorner2(lastColCell)

    _gotoCell(firstColCell)
  }

  const isColumnFullySelected = (col) => {
    const firstColCell = Util.asRef([1, col])
    const lastColCell = Util.asRef([dimensions[0], col])

    return (
      selectedRangeRefs.indexOf(firstColCell) >= 0 &&
      selectedRangeRefs.indexOf(lastColCell) >= 0
    )
  }

  const isRowFullySelected = (row) => {
    const firstRowCell = Util.asRef([row, 1])
    const lastRowCell = Util.asRef([row, dimensions[1]])

    return (
      selectedRangeRefs.indexOf(firstRowCell) >= 0 &&
      selectedRangeRefs.indexOf(lastRowCell) >= 0
    )
  }

  const showExtraRowHeader = dimensions[1] > HEADER_LIMITS.cols
  const showExtraColHeader = dimensions[0] > HEADER_LIMITS.rows

  const highlightedLabelStyle = { backgroundColor: "Gray", color: "White" }

  const colHeaderColumns = (
    <tr style={{ backgroundColor: "lightgray" }}>
      <th onClick={handleSelectAllClick}></th>

      {Util.sequenceMap(dimensions[1], (colBase0) => {
        const col = colBase0 + 1
        const highlightedHeader = isColumnFullySelected(col)
          ? highlightedLabelStyle
          : {}

        return (
          <th
            align="center"
            key={col}
            onClick={handleSelectColClick(col)}
            style={highlightedHeader}
          >
            {Util.colAsLabel(col)}
            {col === dimensions[1] && (
              <>
                {" "}
                <button
                  className="link-button"
                  title="Add new column"
                  onClick={handleAddNewCol}
                >
                  [+]
                </button>
              </>
            )}
          </th>
        )
      })}

      {showExtraRowHeader && <th></th>}
    </tr>
  )

  const rowHeaderColumn = (row) => {
    const defaultStyle = { backgroundColor: "LightGray" }
    const highlightedHeader = isRowFullySelected(row)
      ? highlightedLabelStyle
      : {}

    return (
      <th
        style={{ ...defaultStyle, ...highlightedHeader }}
        onClick={handleSelectRowClick(row)}
      >
        {row}
        {row === dimensions[0] && (
          <>
            <br />
            <button
              className="link-button"
              title="Add new row"
              onClick={handleAddNewRow}
            >
              [+]
            </button>
          </>
        )}
      </th>
    )
  }

  return (
    <>
      <table
        border="1"
        cellSpacing="0"
        cellPadding="4"
        style={{ borderColor: "Silver" }}
      >
        <tbody>
          {colHeaderColumns}

          {Util.sequenceMap(dimensions[0], (rowBase0) => {
            const row = rowBase0 + 1

            return (
              <tr key={row}>
                {rowHeaderColumn(row)}

                {Util.sequenceMap(dimensions[1], (colBase0) => {
                  const col = colBase0 + 1
                  const ref = Util.asRef([row, col])
                  const cell = sheet.findCell([row, col])
                  let cellColor
                  let selectedRangeCellStyle = {}
                  let clipboardCellStyle = {}

                  // https://www.w3schools.com/colors/colors_groups.asp
                  if (cell?.invalid) cellColor = "Salmon"
                  else if (cell?.modified) cellColor = "Gold"
                  else cellColor = "White"

                  if (selectedRangeRefs.indexOf(ref) >= 0)
                    selectedRangeCellStyle = {
                      borderColor: "Blue",
                      backgroundColor: "#d8ecf3",
                      color: "Black"
                    }

                  if (
                    clipboard.range &&
                    Util.expandRange(clipboard.range).flat(2).indexOf(ref) >= 0
                  )
                    clipboardCellStyle = {
                      borderColor: "Blue",
                      color: "Black",
                      borderStyle: { cut: "dashed", copy: "dotted" }[
                        clipboard.action
                      ],
                      borderWidth: "2px"
                    }

                  return (
                    <td
                      valign="top"
                      key={col}
                      style={{
                        backgroundColor: cellColor,
                        ...selectedRangeCellStyle,
                        ...clipboardCellStyle
                      }}
                      onClick={handleCellClick(ref)}
                    >
                      <input
                        type="text"
                        autoComplete="off"
                        size="12"
                        ref={(element) => (cellsRef.current[ref] = element)}
                        defaultValue={cell?.value ?? ""}
                        onKeyDown={handleCellInputKeyDown(ref)}
                        onBlur={handleCellInputBlur(ref)}
                      />
                      <br />
                      <span>
                        {cell && (cell.errorMessage ?? cell.evaluatedValue)}
                      </span>
                    </td>
                  )
                })}

                {showExtraRowHeader && rowHeaderColumn(row)}
              </tr>
            )
          })}

          {showExtraColHeader && colHeaderColumns}
        </tbody>
      </table>

      {DISPLAY_SHEET_JSON && (
        <xmp>{JSON.stringify(sheet, Util.jsonStringifyReplacer, 2)}</xmp>
      )}
    </>
  )
}

Spreadsheet.defaultProps = {}

Spreadsheet.propTypes = {}

export default function App() {
  return <Spreadsheet />
}
