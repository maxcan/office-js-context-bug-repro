
import * as Bluebird from 'bluebird'
import * as _ from 'lodash'
import * as uuid from 'uuid/v4'
import { unstable_renderSubtreeIntoContainer } from 'react-dom';

const MAX_CHUNK_SIZE = 140

export enum Underlines  {
    None = 'None',
    Single = 'Single',
    Double = 'Double',
    Mixed  = 'Mixed'
}
export type Font = Record<'bold' | 'italic', boolean> & {
    highlightColor: string | null
    underline: string | Underlines
}

const rangesForLoc = (context: Word.RequestContext, loc: {paraIdx: number, itemIdx: number, itemCount: number}) => {
    let paragraphs = context.document.body.paragraphs.load('font, styleBuiltIn')
    let paragraph = paragraphs[loc.paraIdx]
    const ranges = paragraph.getTextRanges([' ', ',', '.', ']', ')'], true).slice(loc.itemIdx, loc.itemIdx + loc.itemCount)
    ranges.load('text, font/highlightColor, font/bold, font/underline, font/italic')
    return ranges
}

export const plainFont: Font = {bold: false, italic: false, underline: Underlines.None, highlightColor: null }
const fontTags = ['bold', 'underline', 'italic', 'highlightColor']
const stripFont = (f: Font) =>  _.pick(f, fontTags)
const fontIsEqual = (a: Font | null, b: Font | null) => (a && b && _.isEqual(stripFont(a), stripFont(b)))
const isPlain = (font: Font) => fontIsEqual(font, plainFont)

export class Chunk {
    constructor(range: Word.Range, loc: {paraIdx: number, itemIdx: number}) {
        this.ranges = [range]
        this.style = stripFont(range.font)
        this.loc = {paraIdx: loc.paraIdx, itemIdx: loc.itemIdx, itemCount: 1}
        this.text = range.text
        console.log(`Added chunk ${JSON.stringify(this.loc)}`)
        // this.context = context
        // context.trackedObjects.add(r)
    }
    destroy(ctx: Word.RequestContext) {
        _.each(this.ranges, r => ctx.trackedObjects.remove(r))
    }
    ranges: Array<Word.Range>
    text: string
    loc: {
        paraIdx: number
        itemIdx: number
        itemCount: number
    }
    style: Font
    async addRange (r: Word.Range, _ctx: Word.RequestContext) {
        this.ranges.push(r)
        this.text = this.text + ' ' + r.text
        // _.each(this.ranges, r => ctx.trackedObjects.add(r))
        // const range = this.expandedRange()
        // ctx.load(range!, 'text')
        // ctx.sync().then(() => this.text = range!.text)
    }
    expandedRange() {
        if (this.ranges.length === 0) { return }
        if (this.ranges.length === 1) { return this.ranges[0] }
        return this.ranges[0].expandTo(this.ranges[this.ranges.length - 1])
    }
    async applyStyle(font: Font, ctx: Word.RequestContext) {
        const r = this.expandedRange()
        if (r) { r.font.set(font, {throwOnReadOnly: true}) }
        return ctx.sync()
    }
    async select(ctx: Word.RequestContext) {
        console.log('selecting')
        const ranges = rangesForLoc(ctx, this.loc)
        // if (this.ranges.length === 0) { return }
        // if (this.ranges.length === 1) {
        //     this.ranges[0].select()
        // } else {
            const newRange = (this.ranges[0].expandTo(this.ranges[this.ranges.length - 1]))
            newRange.select()
        // }
        return ctx.sync()
    }
}

export const getChunks = async () => {
    let count = 0
    let controls: Array<Word.ContentControl> = []
    return Word.run(async (context) => {
        try {
            let paragraphs = context.document.body.paragraphs.load('tableNestingLevel, font, styleBuiltIn')
            // let lists = context.document.body.lists.load('levelTypes')
            let wordRanges: Array<Word.RangeCollection> = []
            await context.sync()
            paragraphs.items.forEach(paragraph => {
                const ranges = paragraph.getTextRanges([' ', ',', '.', ']', ')'], true)
                // const ranges = paragraph.getTextRanges, ([' ', ',', ']', '[', '.', '(', ')'], true)
                ranges.load('text, font/highlightColor, font/bold, font/underline, font/italic')
                wordRanges.push(ranges)
            })
            await context.sync()
            type State = { chunk?: Chunk, foundChunks: Array<Chunk>, currentDelim?: string }
            const delims: {} = {
                '[': ']',
                '{': '}',
                '(': ')',
                '\u2018': '\u2019',
                '\u201c': '\u201d'
            }
            let state: State = { foundChunks: [] }
            const updateState = async (nextWord: (Word.Range | null), paraIdx?: number, itemIdx?: number) => {
                if (nextWord) {
                    console.log(`${nextWord.text} data: ${JSON.stringify(stripFont(nextWord.font))}`)
                }
                const saveAndResetCurrentChunk = () => {
                    if (state.chunk) { state.foundChunks.push(state.chunk) }
                    state.currentDelim = undefined
                    state.chunk = undefined
                }
                if (!nextWord || !paraIdx || !itemIdx ) {
                    saveAndResetCurrentChunk()
                    return
                }
                const delimInNextWord = _.find(_.keys(delims), d => nextWord.text.indexOf(d) > -1)
                if (state.chunk && state.currentDelim) {
                    await state.chunk.addRange(nextWord, context)
                    if (nextWord.text.indexOf(delims[state.currentDelim]) > -1 ) {
                        saveAndResetCurrentChunk()
                    }
                } else if (delimInNextWord) {
                    saveAndResetCurrentChunk()
                    state.chunk = new Chunk(nextWord, {paraIdx, itemIdx})
                    state.currentDelim = delimInNextWord
                } else if (isPlain(nextWord.font)) {
                    // found a Plain word with no current delimiter
                    saveAndResetCurrentChunk()
                } else {
                    if (state.chunk) {
                        if (fontIsEqual(state.chunk.style, nextWord.font)) {
                            await state.chunk.addRange(nextWord, context)
                        } else {
                            // in a chunk, but start a new one
                            saveAndResetCurrentChunk()
                            state.chunk = new Chunk(nextWord, {paraIdx, itemIdx})
                        }
                    } else {
                        // start a new chunk
                        state.chunk = new Chunk(nextWord, {paraIdx, itemIdx})
                    }
                }
            }

            wordRanges.forEach((rangesInSingleParagraph, paraIdx)  => {
                rangesInSingleParagraph.items.forEach(async (nextWordRange, itemIdx) => {
                    // console.log(nextWordRange.text, paraIdx, itemIdx)
                    updateState(nextWordRange, paraIdx, itemIdx)
                })
                updateState(null)
            })
            updateState(null)
            return state.foundChunks
        } catch (e) {
            console.error('ERR: ', e)
            console.error(e.trace)
            throw e
        }
    })
}
