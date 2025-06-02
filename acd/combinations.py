import itertools

columns = ["g", "a", "path", "circle", "polyline", "polygon", "text", "points", "style", "d", "character_count"]

combinations_of_columns = []

# Generate all possible combinations of column names
for r in range(1, len(columns) + 1):
    column_combinations = list(itertools.combinations(columns, r))
    combinations_of_columns.extend(column_combinations)

# Convert tuples to lists
combinations_of_columns = [list(comb) for comb in combinations_of_columns]

print(combinations_of_columns)