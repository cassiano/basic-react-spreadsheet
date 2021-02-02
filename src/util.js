const Util = (() => {
  const ALPHABET_LENGTH = "Z".charCodeAt() - "A".charCodeAt() + 1
  const REF_REGEXP = /\b([a-z]+)(\d+)\b/i
  const REF_WITH_ANCHORS_REGEXP = /^([a-z]+)(\d+)$/i
  const RANGE_REGEXP = /\b([a-z]+\d+):([a-z]+\d+)\b/i
  const FORMULA_REGEXP = /^=(.*)$/

  class InvalidRefInFormula {
    constructor(message) {
      this.message = message
      this.name = "InvalidRefInFormula"
    }
  }

  const sequence = (number) => [...Array(number)].map((_, i) => i)
  const sequenceForEach = (number, fn) => sequence(number).forEach((i) => fn(i))
  const sequenceMap = (number, fn) => sequence(number).map((i) => fn(i))
  const sequenceReduce = (number, fn, initialAcc) =>
    sequence(number).reduce(fn, initialAcc)

  const normalizeValue = (value) => {
    if (typeof value === "string") {
      if (value.trim() === "") {
        return ""
      } else {
        return Object.is(Number(value), NaN) ? value.trim() : Number(value)
      }
    } else {
      return value
    }
  }

  // Converts 'A' to 1, 'B' to 2... 'Z' to 26.
  const _colIndexFromSingleLetter = (colSingleRef) => {
    return colSingleRef.charCodeAt() - "A".charCodeAt() + 1
  }

  // Converts 'A' to 1, 'B' to 2... 'Z' to 26, 'AA' to 27 etc
  const colIndexFromLabel = (colRef) => {
    return colRef
      .split("")
      .reduce(
        (acc, letter, i) =>
          acc +
          _colIndexFromSingleLetter(letter) *
            ALPHABET_LENGTH ** (colRef.length - i - 1),
        0
      )
  }

  // Converts 1 to 'A', 2 to 'B'... 26 to 'Z'.
  const _colSingleLetter = (colIndex) => {
    return String.fromCharCode(colIndex - 1 + "A".charCodeAt(0))
  }

  // Converts 1 to 'A', 2 to 'B'... 26 to 'Z', 27 to 'AA' etc
  const colAsLabel = (colIndexOrLabel) => {
    if (typeof colIndexOrLabel === "string") return colIndexOrLabel

    let colIndex = colIndexOrLabel - 1
    let colRef = ""

    while (colIndex >= 0) {
      colRef = _colSingleLetter((colIndex % ALPHABET_LENGTH) + 1) + colRef

      colIndex = Math.trunc(colIndex / ALPHABET_LENGTH) - 1
    }

    return colRef
  }

  const asRef = (refOrCoords) => {
    const { row, col } = asCoords(refOrCoords)

    return colAsLabel(col) + String(row)
  }

  const _rowColFromRef = (ref) => {
    const match = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/i)
    const col = colIndexFromLabel(match[1])
    const row = Number(match[2])

    return { row, col }
  }

  const asCoords = (refOrCoords) => {
    let row, col

    if (refOrCoords instanceof Array) [row, col] = refOrCoords
    else ({ row, col } = _rowColFromRef(refOrCoords))

    if (typeof col === "string") col = colIndexFromLabel(col)

    return { row, col, labelCol: colAsLabel(col) }
  }

  const isFormula = (value) =>
    typeof value === "string" && value.match(FORMULA_REGEXP)

  const rawFormula = (formula) => formula.match(FORMULA_REGEXP)?.[1]

  const findRefsInFormula = (formula) =>
    new Set(textScan(formula.toUpperCase(), REF_REGEXP))

  const templateReplace = (template, replacements) => {
    return Object.entries(replacements).reduce(
      (acc, [ref, value]) =>
        acc.replace(new RegExp(`\\b${ref}\\b`, "gi"), value),
      template
    )
  }

  const templateReplaceForCopy = (template, replacements) => {
    return Object.entries(replacements)
      .reduce(
        (acc, [ref, value]) =>
          acc.replace(
            new RegExp(`\\b${ref.toUpperCase()}\\b`, "g"),
            value.toLowerCase()
          ),
        template.toUpperCase()
      )
      .toUpperCase()
  }

  const templateReplaceForMove = (template, replacements) => {
    return Object.entries(replacements)
      .reduce(
        (acc, [ref, value]) =>
          acc.replace(
            // Ignore ranges.
            new RegExp(`\\b(?<!:)${ref.toUpperCase()}(?!:)\\b`, "g"),
            value.toLowerCase()
          ),
        template.toUpperCase()
      )
      .toUpperCase()
  }

  const textScan = (text, regexp) =>
    [
      ...text.matchAll(
        new RegExp(
          regexp.source,
          [...new Set(regexp.flags.split("")).add("g")].join("")
        )
      )
    ].map((match) => match[0])

  const isEmpty = (value) =>
    value === undefined ||
    value === null ||
    (typeof value === "string" && value.trim() === "")

  const isRef = (ref) => !!ref.match(REF_WITH_ANCHORS_REGEXP)

  const expandRange = (range) => {
    const match = range.match(RANGE_REGEXP)

    const topLeftCoords = Util.asCoords(match[1])
    const bottomRightCoords = Util.asCoords(match[2])

    if (
      bottomRightCoords.row < topLeftCoords.row ||
      bottomRightCoords.col < topLeftCoords.col
    )
      return []

    return Util.sequenceMap(
      bottomRightCoords.row - topLeftCoords.row + 1,
      (rowIndex) =>
        Util.sequenceMap(
          bottomRightCoords.col - topLeftCoords.col + 1,
          (colIndex) =>
            Util.asRef([
              topLeftCoords.row + rowIndex,
              topLeftCoords.col + colIndex
            ])
        )
    )
  }

  const expandRanges = (formula) => {
    const ranges = new Set(textScan(formula, RANGE_REGEXP))

    return [...ranges].reduce(
      (acc, range) =>
        acc.replaceAll(
          range,
          JSON.stringify(expandRange(range)).replaceAll('"', "")
        ),
      formula
    )
  }

  const newRefForCopy = (ref, source, target) => {
    const refCoords = asCoords(ref)
    const sourceCoords = asCoords(source)
    const targetCoords = asCoords(target)

    const newRow = refCoords.row + (targetCoords.row - sourceCoords.row)
    const newCol = refCoords.col + (targetCoords.col - sourceCoords.col)

    if (newRow < 1 || newCol < 1) return "[invalid ref]"

    return asRef([newRow, newCol])
  }

  const setAppend = (targetSet, setOrArrayToAppend) => {
    setOrArrayToAppend.forEach((item) => targetSet.add(item))

    return targetSet
  }

  const setDiff = (oneSet, anotherSet) =>
    new Set([...oneSet].filter((item) => !anotherSet.has(item)))

  const setIntersect = (oneSet, anotherSet) =>
    new Set([...oneSet].filter((item) => anotherSet.has(item)))

  const DIRECTIONS = {
    north: {
      right: "east",
      left: "west",
      walk: (ref, step = 1) => {
        const { row, col } = asCoords(ref)
        return asRef([row > step ? row - step : row, col])
      }
    },
    south: {
      right: "west",
      left: "east",
      walk: (ref, step = 1) => {
        const { row, col } = asCoords(ref)
        return asRef([row + step, col])
      }
    },
    east: {
      right: "south",
      left: "north",
      walk: (ref, step = 1) => {
        const { row, col } = asCoords(ref)
        return asRef([row, col + step])
      }
    },
    west: {
      right: "north",
      left: "south",
      walk: (ref, step = 1) => {
        const { row, col } = asCoords(ref)
        return asRef([row, col > step ? col - step : col])
      }
    }
  }

  function generateSpiralSequence(
    firstSegmentSize,
    initialDirection,
    directionToTurn,
    initialRefsAndValues,
    fn
  ) {
    let currentRef = Object.keys(
      initialRefsAndValues[initialRefsAndValues.length - 1]
    )[0]
    const previousRefs = initialRefsAndValues
      .slice(0, initialRefsAndValues.length - 1)
      .flatMap(Object.keys)
    let direction = initialDirection
    let stepsToWalkInDirection = firstSegmentSize
    let stepsWalkedInDirection = initialRefsAndValues.length

    const count =
      sequenceReduce(firstSegmentSize - 1, (acc, i) => acc + (i + 1), 0) + 1

    return sequenceReduce(
      count - initialRefsAndValues.length,
      (acc, i) => {
        previousRefs.push(currentRef)
        currentRef = DIRECTIONS[direction].walk(currentRef)

        stepsWalkedInDirection++

        if (stepsWalkedInDirection >= stepsToWalkInDirection) {
          direction = DIRECTIONS[direction][directionToTurn]

          stepsToWalkInDirection -= 1
          stepsWalkedInDirection = 1
        }

        acc[currentRef] = fn(
          i,
          previousRefs,
          DIRECTIONS[direction].walk(currentRef)
        )

        return acc
      },
      { ...initialRefsAndValues.reduce((acc, item) => ({ ...acc, ...item })) }
    )
  }

  const generateCellSquares = (side, initialRef, initialValue) => {
    const { row, col } = Util.asCoords(initialRef)

    return Util.sequenceReduce(
      side,
      (acc, i) => ({
        ...acc,
        ..._generateCellColsRows(
          side - i,
          Util.asRef([i + row, i + col]),
          initialValue,
          row,
          col
        )
      }),
      {}
    )
  }

  const _generateCellColsRows = (
    side,
    initialRef,
    initialValue,
    baseRow,
    baseCol
  ) => {
    const { row, col } = Util.asCoords(initialRef)
    const value =
      row === baseRow || col === baseCol
        ? initialValue
        : `=${Util.colAsLabel(col - 1)}${row - 1}+1`

    const verticalResult = Util.sequenceReduce(
      side - 1,
      (acc, i) => {
        const key = `${Util.colAsLabel(col)}${i + 1 + row}`
        const value = `=${Util.colAsLabel(col)}${i + row}+1`

        acc[key] = value

        return acc
      },
      { [initialRef]: value }
    )

    const horizontalResult = Util.sequenceReduce(
      side - 1,
      (acc, i) => {
        const key = `${Util.colAsLabel(i + col + 1)}${row}`
        const value = `=${Util.colAsLabel(i + col)}${row}+1`

        acc[key] = value

        return acc
      },
      { [initialRef]: value }
    )

    return { ...verticalResult, ...horizontalResult }
  }

  const showElapsedTimes = (
    fn,
    { message = "(no description)", repeat = 1, matchResults = true } = {}
  ) => {
    const runs = sequenceMap(repeat, (i) => {
      const [startTime, fnResult, endTime] = [Date.now(), fn(), Date.now()]

      return { result: fnResult, elapsedTime: endTime - startTime }
    })

    if (matchResults) {
      const resultsAsJson = runs.map(({ result }) =>
        JSON.stringify(result, jsonStringifyReplacer)
      )

      if (resultsAsJson.slice(1).some((result) => result !== resultsAsJson[0]))
        throw new Error(
          `Results don't match for '${message}'. Please supply a pure function!`
        )
    }

    const elapsedTimes = runs.map(({ elapsedTime }) => elapsedTime)
    const minElapsedTime = Math.min(...elapsedTimes)
    const maxElapsedTime = Math.max(...elapsedTimes)
    const AvgElapsedTime =
      elapsedTimes.reduce((acc, time) => acc + time) / elapsedTimes.length

    console.log(`Elapsed times for '${message}' after ${repeat} runs:`, {
      minElapsedTime,
      maxElapsedTime,
      AvgElapsedTime
    })

    return runs[0].result
  }

  function jsonStringifyReplacer(key, value) {
    if (value instanceof Set) {
      return [...value]
    } else if (key === "sheet") {
      return `(Sheet ${value.id})`
    } else if (value === undefined) {
      return "(undefined)"
    } else if (value === null) {
      return "(null)"
    } else {
      return value
    }
  }

  return {
    RANGE_REGEXP,
    InvalidRefInFormula,
    sequence,
    sequenceForEach,
    sequenceMap,
    sequenceReduce,
    normalizeValue,
    colIndexFromLabel,
    colAsLabel,
    asRef,
    asCoords,
    isFormula,
    rawFormula,
    findRefsInFormula,
    isEmpty,
    templateReplace,
    templateReplaceForCopy,
    templateReplaceForMove,
    textScan,
    isRef,
    expandRange,
    expandRanges,
    newRefForCopy,
    setAppend,
    setDiff,
    setIntersect,
    generateSpiralSequence,
    generateCellSquares,
    showElapsedTimes,
    jsonStringifyReplacer
  }
})()

export default Util
