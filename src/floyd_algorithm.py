"""
CS 232/AI 209 - Discrete Structures in Computer Science
Spring 2026 - Prof. Ping-Tsai Chung
Mini-Project: Floyd's Algorithm for All-Pairs Shortest Paths
Due: May 4, 2026
"""

INF = float('inf')


def floyd_warshall(graph, city_names):
    """
    Floyd's Algorithm for All-Pairs Shortest Paths.
    Data Structure: 2D adjacency/distance matrix
    Time Complexity:  O(n^3)
    Space Complexity: O(n^2)
    """
    n = len(graph)
    dist = [row[:] for row in graph]
    next_node = [[None] * n for _ in range(n)]

    for i in range(n):
        for j in range(n):
            if i != j and dist[i][j] != INF:
                next_node[i][j] = j

    print("=" * 70)
    print("FLOYD'S ALGORITHM - ALL PAIRS SHORTEST PATHS")
    print("=" * 70)
    print(f"Vertices: {', '.join(city_names)}\n")
    print_matrix(dist, city_names, "D^(0) - Initial Weight Matrix")

    for k in range(n):
        for i in range(n):
            for j in range(n):
                if dist[i][k] != INF and dist[k][j] != INF:
                    if dist[i][k] + dist[k][j] < dist[i][j]:
                        dist[i][j] = dist[i][k] + dist[k][j]
                        next_node[i][j] = next_node[i][k]
        print_matrix(dist, city_names,
                     f"D^({k+1}) - intermediate: {city_names[k]}")

    return dist, next_node


def print_matrix(matrix, names, title):
    n = len(matrix)
    col_w = max(len(nm) for nm in names) + 2
    val_w = 9
    print(f"\n{title}")
    print("-" * (col_w + val_w * n))
    print(" " * col_w + "".join(f"{nm:>{val_w}}" for nm in names))
    print("-" * (col_w + val_w * n))
    for i, rn in enumerate(names):
        row = f"{rn:<{col_w}}"
        for j in range(n):
            v = matrix[i][j]
            row += f"{'INF':>{val_w}}" if v == INF else f"{v:>{val_w}}"
        print(row)
    print()


def reconstruct_path(next_node, start, end, city_names):
    if next_node[start][end] is None:
        return None
    route = [city_names[start]]
    cur = start
    seen = set()
    while cur != end:
        if cur in seen:
            return None
        seen.add(cur)
        cur = next_node[cur][end]
        route.append(city_names[cur])
    return route


def print_all_paths(dist, next_node, city_names):
    n = len(city_names)
    print("\n" + "=" * 70)
    print("ALL SHORTEST PATHS")
    print("=" * 70)
    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            if dist[i][j] == INF:
                print(f"{city_names[i]} -> {city_names[j]}: NO PATH")
            else:
                route = reconstruct_path(next_node, i, j, city_names)
                path_str = " -> ".join(route) if route else "direct"
                print(f"{city_names[i]} -> {city_names[j]}: "
                      f"Length={dist[i][j]},  Path: {path_str}")


def test_lecture_example():
    print("\n\nTEST 1: LECTURE EXAMPLE (Slides Fig 8.16)")
    names = ["a", "b", "c", "d"]
    I = INF
    graph = [
        [0, I, 3, I],
        [2, 0, I, I],
        [I, 7, 0, 1],
        [6, I, I, 0],
    ]
    dist, nxt = floyd_warshall(graph, names)
    print_all_paths(dist, nxt, names)


def test_us_cities():
    print("\n\nTEST 2: US CITIES GRAPH (Figure 2)")
    names = [
        "Seattle", "SanFrancisco", "LosAngeles",
        "Denver", "KansasCity", "Dallas",
        "Houston", "Chicago", "Boston",
        "NewYork", "Atlanta", "Miami"
    ]
    I = INF
    graph = [
        [0,    807,  I,    1331, I,    I,    I,    2097, I,    I,    I,    I   ],
        [807,  0,    381,  1267, 1015, I,    I,    I,    I,    I,    I,    I   ],
        [I,    381,  0,    I,    1663, 1435, I,    I,    I,    I,    I,    I   ],
        [1331, 1267, I,    0,    599,  I,    I,    1003, I,    I,    I,    I   ],
        [I,    1015, 1663, 599,  0,    496,  I,    533,  I,    I,    864,  I   ],
        [I,    I,    1435, I,    496,  0,    239,  I,    I,    I,    781,  I   ],
        [I,    I,    I,    I,    I,    239,  0,    I,    I,    I,    810,  1187],
        [2097, I,    I,    1003, 533,  I,    I,    0,    983,  787,  I,    I   ],
        [I,    I,    I,    I,    I,    I,    I,    983,  0,    214,  I,    I   ],
        [I,    I,    I,    I,    I,    I,    I,    787,  214,  0,    1260, I   ],
        [I,    I,    I,    I,    864,  781,  810,  I,    I,    1260, 0,    661 ],
        [I,    I,    I,    I,    I,    I,    1187, I,    I,    I,    661,  0   ],
    ]
    dist, nxt = floyd_warshall(graph, names)
    print_all_paths(dist, nxt, names)


if __name__ == "__main__":
    test_lecture_example()
    test_us_cities()
