import * as React from 'react'
import { Container, ListGroup, ListGroupItem, Button, Label, Input, ButtonGroup, Row } from 'reactstrap'

class Chunk {
    range: Word.Range
    text: string
    // context: Word.RequestContext
    constructor(t: string, r: Word.Range, _ctx: Word.RequestContext) {
        r.track()
        this.range = r
        this.text = t
        // this.range.track()
        // ctx.trackedObjects.add(r)
        // r.track()
        // this.context = ctx
    }
}

const getChunks = async () => {
    return await Word.run(async context => {
        let paragraphs = context.document.body.paragraphs.load()
        let wordRanges: Array<Word.RangeCollection> = []
        await context.sync()
        paragraphs.track()
        paragraphs.items.forEach(paragraph => {
            const ranges = paragraph.getTextRanges([' ', ',', '.', ']', ')'], true)
            ranges.load('text')
            wordRanges.push(ranges)
        })
        await context.sync()
        let chunks: Chunk[] = []
        wordRanges.forEach(ranges => ranges.items.forEach(range => {
            context.trace('added range')
            chunks.push(new Chunk(range.text, range, context))
            console.log('tracking')

        }))
        await context.sync()
        return chunks
    })

}

interface ChunkControlProps { chunk: Chunk; onSelect: (e: React.MouseEvent<HTMLElement>) => void }
export const ChunkControl: React.SFC<ChunkControlProps> = ({ chunk, onSelect}) => {
    return (
        <div style={{marginLeft: '0.5em'}}><a href='#' onClick={onSelect}>{chunk.text}</a></div>
    )
}
declare var OfficeExtension: any;

export class App extends React.Component<{title: string}, {chunks: Chunk[]}> {
    constructor(props, context) {
        super(props, context)
        this.state = { chunks: [] }
    }

    componentDidMount() { this.click() }

    click = async () => {
        const chunks = await getChunks()
        this.setState(prev => ({ ...prev, chunks: chunks }))
    }

    onSelectRange(chunk: Chunk) {
        return async (e: React.MouseEvent<HTMLElement>) => {
            e.preventDefault()
            console.log('about to run word.run')
            await Word.run(chunk.range, async ctx => {
                console.log('about to select')
                chunk.range.select();
                console.log('about to sync')
                await ctx.sync().catch(e => {
                    console.error(e.stack)
                    console.error(e.debugInfo)
                })
                console.log('synced')
            })
            console.log('just ran')
        }
    }

    render() {
        return (
            <Container fluid={true}>
                <Button color='primary' size='sm' block className='ms-welcome__action' onClick={this.click}>Find Chunks demo</Button>
                <hr/>
                <ListGroup>
                    {this.state.chunks.map((chunk, idx) => (
                        <ListGroupItem key={idx}>
                            <ChunkControl  onSelect={this.onSelectRange(chunk)} chunk={chunk}/>
                        </ListGroupItem>
                    ))}
                </ListGroup>
            </Container>
        )
    };
};
