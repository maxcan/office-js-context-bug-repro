import * as React from 'react'
import { Container, ListGroup, ListGroupItem, Button, Label, Input, ButtonGroup, Row } from 'reactstrap'

const getChunks = async () => {
    return await Word.run(async context => {
        let paragraphs = context.document.body.paragraphs.load()
        let wordRanges: Array<Word.RangeCollection> = []
        await context.sync()
        paragraphs.track()
        paragraphs.items.forEach(paragraph => {
            const ranges = paragraph.getTextRanges([' ', ',', '.', ']', ')'], true)
            paragraph.track();
            ranges.track();
            ranges.load('text')
            wordRanges.push(ranges)
        })
        await context.sync()
        let chunks: Word.Range[] = []
        wordRanges.forEach(ranges => ranges.items.forEach(range => {
            range.track()
            chunks.push(range)
        }))
        await context.sync()
        return chunks
    })

}

interface ChunkControlProps { range: Word.Range; onSelect: (e: React.MouseEvent<HTMLElement>) => void }
export const ChunkControl: React.SFC<ChunkControlProps> = ({ range, onSelect}) => {
    return (

        <div style={{marginLeft: '0.5em'}}><a href='#' onClick={onSelect}>{range.text}</a></div>
    )
}

export class App extends React.Component<{title: string}, {chunks: Word.Range[]}> {
    constructor(props, context) {
        super(props, context)
        this.state = { chunks: [] }
    }

    componentDidMount() { this.click() }

    click = async () => {
        const chunks = await getChunks()
        this.setState(prev => ({ ...prev, chunks: chunks }))
    }

    onSelectRange(range: Word.Range) {
        return async (e: React.MouseEvent<HTMLElement>) => {
            e.preventDefault()
            await Word.run(range, async _ctx => {
                range.font.set({bold: true})
                range.select()
                // range.font.set({bold: true})
                // await ctx.sync().catch(e => {
                //     console.error(e.stack)
                //     console.error(e.debugInfo)
                // })
            })
        }
    }

    render() {
        return (
            <Container fluid={true}>
                <Button color='primary' size='sm' block className='ms-welcome__action' onClick={this.click}>Find Chunks </Button>
                <hr/>
                <ListGroup>
                    {this.state.chunks.map((chunk, idx) => (
                        <ListGroupItem key={idx}>
                            <ChunkControl  onSelect={this.onSelectRange(chunk)} range={chunk}/>
                        </ListGroupItem>
                    ))}
                </ListGroup>
            </Container>
        )
    };
};
