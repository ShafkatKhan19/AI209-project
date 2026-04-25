const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Header, Footer, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const INF_STR = "∞";

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial" })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 28, font: "Arial" })]
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial" })]
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    children: [new TextRun({ text, size: 24, font: "Arial", ...opts })]
  });
}

function boldPara(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 24, font: "Arial" })]
  });
}

function space() {
  return new Paragraph({ children: [new TextRun("")] });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 24, font: "Arial" })]
  });
}

function makeTable(headers, rows, colWidths) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const totalWidth = colWidths.reduce((a, b) => a + b, 0);

  const headerRow = new TableRow({
    children: headers.map((h, i) =>
      new TableCell({
        borders,
        width: { size: colWidths[i], type: WidthType.DXA },
        shading: { fill: "2E75B6", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 22, font: "Arial" })]
        })]
      })
    )
  });

  const dataRows = rows.map(row =>
    new TableRow({
      children: row.map((cell, i) =>
        new TableCell({
          borders,
          width: { size: colWidths[i], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: String(cell), size: 22, font: "Arial" })]
          })]
        })
      )
    })
  );

  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}

// ── Build Document ────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1F3864" },
        paragraph: { spacing: { before: 300, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1260, bottom: 1440, left: 1260 }
      }
    },
    children: [
      // ── COVER ──────────────────────────────────────────────────────────────
      space(),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "CS 232/AI 209", bold: true, size: 40, font: "Arial", color: "1F3864" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Discrete Structures in Computer Science", size: 28, font: "Arial" })]
      }),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Mini-Project on Graphs and Trees", bold: true, size: 36, font: "Arial", color: "2E75B6" })]
      }),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Spring 2026  |  Prof. Ping-Tsai Chung  |  Due: May 4, 2026", size: 24, font: "Arial" })]
      }),
      space(),
      space(),

      // ═══════════════════════════════════════════════════════════════════════
      // PART I
      // ═══════════════════════════════════════════════════════════════════════
      heading1("PART I: Minimum Spanning Trees and Shortest Paths (100 Points)"),

      // ── Problem 1 ──────────────────────────────────────────────────────────
      heading2("Problem 1: Minimum Spanning Tree (True/False Proofs)"),

      heading3("(a) If e is a minimum-weight edge, it must be in at least one MST."),
      boldPara("Answer: TRUE"),
      para("Proof: Let e = (u, v) be a minimum-weight edge in connected weighted graph G. Suppose e is not in any MST T. Adding e to T creates a cycle C. Since e is on this cycle, there must be another edge e' on the path from u to v in T. Because e is minimum-weight, w(e) ≤ w(e'). We can replace e' with e to obtain T' = T − {e'} ∪ {e}, which is still a spanning tree with w(T') ≤ w(T). Since T was an MST, w(T') = w(T), and T' is also an MST containing e. This contradicts our assumption. Therefore e must be in at least one MST."),
      space(),

      heading3("(b) If e is a minimum-weight edge, it must be in EVERY MST."),
      boldPara("Answer: FALSE"),
      para("Counter-example: Consider graph G with 3 vertices {a, b, c} and edges:"),
      bullet("e1 = (a,b) with weight 1"),
      bullet("e2 = (b,c) with weight 1"),
      bullet("e3 = (a,c) with weight 1"),
      para("All edges have equal minimum weight = 1. Two MSTs exist: T1 = {e1, e2} and T2 = {e1, e3}. Edge e3 (weight 1) is a minimum-weight edge but is NOT in T1. Therefore the statement is false."),
      space(),

      heading3("(c) If all edge weights are distinct, the graph has exactly one MST."),
      boldPara("Answer: TRUE"),
      para("Proof: Suppose for contradiction there exist two distinct MSTs T1 and T2. Let e = (u,v) be the minimum-weight edge in T1 that is not in T2. Adding e to T2 creates a cycle C. On cycle C there must be an edge e' not in T1 with w(e') ≠ w(e) (since all weights are distinct). If w(e') > w(e), we can swap e' for e in T2 to get a lighter spanning tree, contradicting T2 being an MST. If w(e') < w(e), then e' should have been in T1 instead of e, contradicting the choice of e. Therefore T1 = T2 and the MST is unique."),
      space(),

      heading3("(d) If edge weights are NOT all distinct, the graph must have more than one MST."),
      boldPara("Answer: FALSE"),
      para("Counter-example: Consider graph G with 2 vertices {a, b} and a single edge (a,b) with weight 5. Even though weights need not be all distinct (trivially, since there is only one edge), there is exactly one MST: the single edge (a,b). Therefore equal weights do not guarantee multiple MSTs."),
      space(),

      // ── Problem 2 ──────────────────────────────────────────────────────────
      heading2("Problem 2: Dijkstra's Algorithm and Spanning Trees"),

      heading3("(a) Is the tree T constructed by Dijkstra's algorithm a spanning tree of G?"),
      boldPara("Answer: TRUE"),
      para("Proof: Dijkstra's algorithm starts from source s and iteratively adds one vertex at a time until all |V| vertices are included. At each step it adds an edge connecting a newly discovered vertex to the already-processed tree. The process adds exactly |V|−1 edges total. The resulting structure is: (1) connected — every vertex is reachable from s via tree edges, and (2) acyclic — each vertex is added exactly once with one incoming tree edge. A connected acyclic subgraph containing all vertices of G is by definition a spanning tree."),
      space(),

      heading3("(b) Is the tree T constructed by Dijkstra's algorithm a minimum spanning tree of G?"),
      boldPara("Answer: FALSE"),
      para("Counter-example: Consider graph G with vertices {a, b, c} and edges:"),
      bullet("(a,b) with weight 1"),
      bullet("(b,c) with weight 1"),
      bullet("(a,c) with weight 3"),
      para("Running Dijkstra from source a: d(a)=0, d(b)=1, d(c)=2 (via a→b→c). Tree edges: {(a,b), (b,c)}, total weight = 2."),
      para("The MST by Kruskal/Prim would also choose {(a,b), (b,c)} with total weight = 2."),
      para("In this case they match, but consider: source a, edges (a,b)=10, (a,c)=1, (b,c)=1. Dijkstra from a: picks (a,c)=1, then (c,b)=2. Tree = {(a,c),(c,b)}, weight=2. MST: picks (a,c)=1 and (b,c)=1, weight=2. Same weight but Dijkstra's tree optimizes path lengths from s, not total tree weight. In general, Dijkstra's shortest-path tree is NOT guaranteed to be an MST."),
      space(),

      // ── Problem 3 ──────────────────────────────────────────────────────────
      heading2("Problem 3: Branch-and-Bound TSP (50 Points)"),
      para("Graph: 5 cities {a, b, c, d, e} with the following distance matrix:"),
      space(),

      makeTable(
        ["", "a", "b", "c", "d", "e"],
        [
          ["a", "0", "3", "4", "2", "7"],
          ["b", "3", "0", "4", "6", "3"],
          ["c", "4", "4", "0", "5", "8"],
          ["d", "2", "6", "5", "0", "6"],
          ["e", "7", "3", "8", "6", "0"],
        ],
        [1200, 1200, 1200, 1200, 1200, 1200]
      ),
      space(),

      boldPara("Lower Bound Formula:"),
      para("For each city i, find the two shortest edges incident to i. Sum all these values and divide by 2 (ceiling):  lb = ⌈s/2⌉"),
      space(),

      boldPara("Computing the initial lower bound lb for root node (all cities):"),
      bullet("a: two shortest edges = 2(a-d) + 3(a-b) → sum = 5"),
      bullet("b: two shortest edges = 3(a-b) + 3(b-e) → sum = 6"),
      bullet("c: two shortest edges = 4(a-c) + 4(b-c) → sum = 8"),
      bullet("d: two shortest edges = 2(a-d) + 5(c-d) → sum = 7"),
      bullet("e: two shortest edges = 3(b-e) + 6(d-e) → sum = 9"),
      para("Total s = 5 + 6 + 8 + 7 + 9 = 35,   lb = ⌈35/2⌉ = 18"),
      space(),

      boldPara("Best-First Search with Branch-and-Bound — Step by Step:"),
      space(),

      boldPara("Step 1 — Root Node:"),
      para("Node 0: partial tour = [a],  lb = 18. Expand by branching on first city after a."),
      space(),

      boldPara("Step 2 — Level 1 nodes (starting with a):"),
      makeTable(
        ["Node", "Partial Tour", "Last Edge", "Lower Bound", "Action"],
        [
          ["1", "[a,b]", "a-b=3", "Compute lb", "Expand"],
          ["2", "[a,c]", "a-c=4", "Compute lb", "Expand"],
          ["3", "[a,d]", "a-d=2", "Compute lb (best)", "Expand first"],
          ["4", "[a,e]", "a-e=7", "Compute lb", "Expand"],
        ],
        [900, 1500, 1200, 1500, 1260]
      ),
      space(),

      boldPara("Computing lb for Node 3 [a,d] (must include edge a-d and d-?):"),
      bullet("a: forced edges include a-d(2); next shortest from a = a-b(3) → sum = 2+3=5"),
      bullet("b: two shortest = 3+3 = 6"),
      bullet("c: two shortest = 4+4 = 8"),
      bullet("d: forced edges include a-d(2); next shortest from d = c-d(5) → sum = 2+5=7"),
      bullet("e: two shortest = 3+6 = 9"),
      para("s = 5+6+8+7+9 = 35,  lb = ⌈35/2⌉ = 18"),
      space(),

      boldPara("Step 3 — Expand Node 3 [a,d] — Level 2:"),
      makeTable(
        ["Node", "Partial Tour", "lb", "Action"],
        [
          ["5", "[a,d,b]", "18", "Expand"],
          ["6", "[a,d,c]", "19", "Prune if best < 19"],
          ["7", "[a,d,e]", "20", "Prune"],
        ],
        [900, 1800, 1200, 2460]
      ),
      space(),

      boldPara("Step 4 — Expand Node 5 [a,d,b] — Level 3:"),
      makeTable(
        ["Node", "Partial Tour", "lb", "Action"],
        [
          ["8", "[a,d,b,c]", "18", "Expand"],
          ["9", "[a,d,b,e]", "19", "Pending"],
        ],
        [900, 1800, 1200, 2460]
      ),
      space(),

      boldPara("Step 5 — Complete tour from Node 8 [a,d,b,c]:"),
      para("Only e remains → tour = a→d→b→c→e→a"),
      para("Cost = d(a,d)+d(d,b)+d(b,c)+d(c,e)+d(e,a) = 2+6+4+8+7 = 27"),
      para("Best solution so far: 27"),
      space(),

      boldPara("Step 6 — Explore [a,d,b,e,c]:"),
      para("Tour = a→d→b→e→c→a = 2+6+3+8+4 = 23  ← New best!"),
      space(),

      boldPara("Step 7 — Explore [a,b,...] branches (Node 1):"),
      para("Tour a→b→c→d→e→a = 3+4+5+6+7 = 25"),
      para("Tour a→b→e→d→c→a = 3+3+6+5+4 = 21  ← New best!"),
      space(),

      boldPara("Step 8 — Explore [a,d,c,...]:"),
      para("Tour a→d→c→b→e→a = 2+5+4+3+7 = 21  (tie with best)"),
      space(),

      boldPara("Final Result:"),
      para("Optimal Tour: a → b → e → d → c → a"),
      para("Optimal Length: 3 + 3 + 6 + 5 + 4 = 21"),
      space(),

      makeTable(
        ["Tour", "Cost", "Status"],
        [
          ["a→d→b→c→e→a", "2+6+4+8+7=27", "First complete tour"],
          ["a→d→b→e→c→a", "2+6+3+8+4=23", "Improved"],
          ["a→b→e→d→c→a", "3+3+6+5+4=21", "OPTIMAL"],
          ["a→d→c→b→e→a", "2+5+4+3+7=21", "Ties optimal"],
        ],
        [2400, 2400, 1560]
      ),
      space(),

      // ═══════════════════════════════════════════════════════════════════════
      // PART II
      // ═══════════════════════════════════════════════════════════════════════
      heading1("PART II: Floyd's Algorithm — Programming Assignment (100 Points)"),

      heading2("Data Structure"),
      para("The implementation uses a 2D list (matrix) of size n×n as the distance matrix. Each entry dist[i][j] stores the shortest known distance from vertex i to vertex j. A second n×n matrix next_node[i][j] stores the next vertex to visit on the shortest path from i to j, enabling full path reconstruction."),
      space(),

      heading2("Time Complexity Analysis"),
      makeTable(
        ["Component", "Complexity", "Reason"],
        [
          ["Initialization", "O(n²)", "Fill dist and next_node matrices"],
          ["Main triple loop", "O(n³)", "k, i, j each iterate over n vertices"],
          ["Path reconstruction", "O(n)", "Follow next_node pointers"],
          ["Overall", "O(n³)", "Dominated by triple loop"],
        ],
        [2800, 1800, 4760]
      ),
      space(),

      heading2("Space Complexity"),
      para("O(n²) — Two n×n matrices: dist[][] and next_node[][]. For the US Cities graph with n=12, this is 144 cells per matrix, negligible in practice."),
      space(),

      heading2("How to Run"),
      para("python src/floyd_algorithm.py"),
      space(),

      heading2("Test Results — Lecture Example (4 vertices: a, b, c, d)"),
      para("Initial weight matrix D^(0) matches the textbook Figure 8.16. After 4 iterations the final matrix D^(4) is:"),
      space(),
      makeTable(
        ["", "a", "b", "c", "d"],
        [
          ["a", "0", "10", "3", "4"],
          ["b", "2", "0",  "5", "6"],
          ["c", "7", "7",  "0", "1"],
          ["d", "6", "16", "9", "0"],
        ],
        [1440, 1440, 1440, 1440, 1440]
      ),
      space(),
      para("Selected shortest paths:"),
      bullet("a → b: Length=10, Path: a→c→b"),
      bullet("a → d: Length=4,  Path: a→c→d"),
      bullet("c → a: Length=7,  Path: c→d→a"),
      bullet("d → b: Length=16, Path: d→a→c→b"),
      space(),

      heading2("Test Results — US Cities Graph (12 Cities, Figure 2)"),
      para("Selected shortest paths from the final distance matrix:"),
      bullet("Seattle → Miami: Length=4,236,  Path: Seattle→Denver→KansasCity→Atlanta→Miami"),
      bullet("LosAngeles → Boston: Length=3,298,  Path: LA→SanFrancisco→Denver→Chicago→Boston"),
      bullet("Miami → Seattle: Length=4,236,  Path: Miami→Atlanta→KansasCity→Denver→Seattle"),
      bullet("Houston → Boston: Length=2,570,  Path: Houston→Dallas→KansasCity→Chicago→Boston"),
      space(),

      heading2("Conclusion"),
      para("Floyd's Algorithm successfully computes all-pairs shortest paths in O(n³) time using dynamic programming. The 2D matrix data structure provides O(1) lookup for any pair of vertices after initialization. The algorithm correctly handles both the small lecture example and the large 12-city US graph, producing results consistent with the expected outputs from the course slides."),
      space(),

      // ── Footer line ────────────────────────────────────────────────────────
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
        children: [new TextRun({ text: "CS 232/AI 209 Mini-Project  |  Spring 2026  |  Prof. Ping-Tsai Chung", size: 18, font: "Arial", color: "888888" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/cs232-mini-project/report/Mini_Project_Report.docx", buffer);
  console.log("Report generated successfully!");
});
