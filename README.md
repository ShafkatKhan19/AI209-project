# CS 232/AI 209 - Mini-Project: Graphs and Trees
**Spring 2026 | Prof. Ping-Tsai Chung | Due: May 4, 2026**

## Overview
This project covers two parts:
- **Part I**: Theory — Minimum Spanning Trees, Dijkstra's Algorithm, and Branch-and-Bound TSP
- **Part II**: Programming — Floyd's Algorithm for All-Pairs Shortest Paths

## Repository Structure
```
cs232-mini-project/
├── src/
│   └── floyd_algorithm.py     # Floyd's Algorithm implementation
├── report/
│   └── Mini_Project_Report.docx  # Full written report with all solutions
└── README.md
```

## How to Run
```bash
python src/floyd_algorithm.py
```

### Requirements
- Python 3.x (no external libraries needed)

## What the Program Does
- Implements **Floyd's Algorithm** O(n³) for all-pairs shortest paths
- Prints each iteration matrix D^(0) through D^(n)
- Reconstructs and prints the actual shortest path between every pair of vertices
- **Test 1**: Lecture slide example (4 vertices)
- **Test 2**: US Cities graph (12 cities, Figure 2 from assignment)

## Algorithm Summary
| Property | Value |
|----------|-------|
| Data Structure | 2D Distance Matrix (n × n) |
| Time Complexity | O(n³) |
| Space Complexity | O(n²) |
| Approach | Dynamic Programming |

## Part I Summary
- **Problem 1a**: TRUE — minimum-weight edge is in at least one MST (Cut Property)
- **Problem 1b**: FALSE — not necessarily in EVERY MST (counter-example provided)
- **Problem 1c**: TRUE — distinct weights → unique MST
- **Problem 1d**: FALSE — equal weights don't guarantee multiple MSTs
- **Problem 2a**: TRUE — Dijkstra's tree is a spanning tree
- **Problem 2b**: FALSE — it is a shortest-path tree, not necessarily an MST
- **Problem 3**: Branch-and-Bound TSP on 5-city graph (see report for full step-by-step)
