import { useState } from 'react'
import { Play, Eraser, TerminalSquare } from 'lucide-react'
import { Lexer } from './compiler/lexer'
import { Parser } from './compiler/parser'
import { Evaluator } from './compiler/evaluator'

function App() {
  const defaultSnippet = `Function AddNumbers(a, b)
  AddNumbers = a + b
End Function

Sub MainLoop()
  for i = 1 to 3
    debug.print AddNumbers(i, 10)
  next i
End Sub

MainLoop`
  const [code, setCode] = useState(defaultSnippet)
  const [output, setOutput] = useState<string[]>([])

  const handleRun = () => {
    setOutput(['> Compiling and running...'])

    try {
      const lexer = new Lexer(code)
      const tokens = lexer.tokenize()

      const parser = new Parser(tokens)
      const ast = parser.parse()

      const newOutputs: string[] = []
      const evaluator = new Evaluator((out) => {
        newOutputs.push(out)
      })

      evaluator.evaluate(ast)

      setOutput(prev => [...prev, ...newOutputs, '> Execution finished.'])
    } catch (err: any) {
      setOutput(prev => [...prev, `Error: ${err.message}`])
    }
  }

  const handleClear = () => {
    setOutput([])
  }

  return (
    <div className="flex flex-col h-screen bg-neutral-950 text-neutral-100 font-sans">
      {/* Header */}
      <header className="flex items-center justify-between px-6 py-4 bg-neutral-900 border-b border-neutral-800 shadow-sm">
        <div className="flex items-center gap-3">
          <TerminalSquare className="w-6 h-6 text-blue-500" />
          <h1 className="text-xl font-semibold tracking-tight">VBA Web Compiler</h1>
        </div>
        <div className="flex items-center gap-3">
          <button
            onClick={handleClear}
            className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-neutral-400 bg-neutral-800 rounded-md hover:bg-neutral-700 hover:text-white transition-colors cursor-pointer"
          >
            <Eraser className="w-4 h-4" />
            Clear
          </button>
          <button
            onClick={handleRun}
            className="flex items-center gap-2 px-6 py-2 text-sm font-semibold text-white bg-blue-600 rounded-md hover:bg-blue-500 shadow-md shadow-blue-500/20 transition-all cursor-pointer"
          >
            <Play className="w-4 h-4" />
            Run
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex flex-1 overflow-hidden">
        {/* Editor Pane */}
        <section className="flex-1 flex flex-col border-r border-neutral-800 relative">
          <div className="px-4 py-2 bg-neutral-900 border-b border-neutral-800 text-xs font-mono text-neutral-400 flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-green-500"></span>
            main.vba
          </div>
          <textarea
            value={code}
            onChange={(e) => setCode(e.target.value)}
            spellCheck={false}
            className="flex-1 w-full bg-neutral-950 text-neutral-200 font-mono text-sm sm:text-base p-4 focus:outline-none resize-none leading-relaxed"
            placeholder="Type your VBA code here..."
          />
        </section>

        {/* Console Pane */}
        <section className="flex-1 flex flex-col bg-black">
          <div className="px-4 py-2 bg-neutral-900 border-b border-neutral-800 text-xs font-mono text-neutral-400 flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-blue-500"></span>
            Console Output
          </div>
          <div className="flex-1 p-4 overflow-y-auto font-mono text-sm text-neutral-300">
            {output.length === 0 ? (
              <span className="text-neutral-600 italic">Ready...</span>
            ) : (
              output.map((line, i) => (
                <div key={i} className="whitespace-pre-wrap leading-relaxed">{line}</div>
              ))
            )}
          </div>
        </section>
      </main>
    </div>
  )
}

export default App
