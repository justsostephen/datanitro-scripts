"""Calculate Lumi Show quiz scores."""

from operator import itemgetter

# Define correct answers.
correct = [3, 2, 1, 4, 1, 2, 4, 3, 4, 2, 1, 3]

# Read table details.
active_sheet('discussions')
tables = []
# Set data start row.
row = 5
# Loop through rows until empty row is found.
data_present = True
while data_present:
    # Append nested list to 'tables' list containing delegate ID and table.
    tables.append([Cell(row, 'E').value, Cell(row, 'F').value])
    row += 1
    # Check next row for data.
    data_present = Cell(row, 'A').value

# Read delegate details.
active_sheet('questions')
delegates = []
# Set data start row.
row = 4
# Determine first quiz response column.
header = Cell(3, 'A').horizontal
for i in range(0, len(header)):
    if header[i][-12:] == ': Question 1':
        quiz_col = i + 1
        break
# Loop through rows until empty row is found.
data_present = True
while data_present:
    # Read delegate ID.
    delegate_id = Cell(row, 'C').value
    # If the delegate has entered a table, record their ID, name, answers,
    # initialised score and table.
    for table_id, table in tables:
        if table_id == delegate_id:
            delegates.append([delegate_id, Cell(row, 'B').value,
                Cell(row, quiz_col).horizontal, 0, table])
            break
    row += 1
    # Check next row for data.
    data_present = Cell(row, 'A').value

# Determine the number of questions answered based on header row.
answered = len(Cell(3, quiz_col).horizontal)
# Calculate delegate scores. Add 5 to score for a correct answer, otherwise
# subtract 5.
for delegate in delegates:
    for i in range(0, len(correct[:answered])):
        if delegate[2][i] == correct[i]:
            delegate[3] += 5
        else:
            delegate[3] -= 3

# Sort 'delegates' by ascending table then descending score.
delegates.sort(key=itemgetter(4))
delegates.sort(key=itemgetter(3), reverse=True)

# Write delegate scores.
active_sheet('scores')
Cell(1, 'A').value = 'Name'
Cell(1, 'B').value = 'Table'
Cell(1, 'C').value = 'Score'
row = 2
for delegate_id, name, answers, score, table in delegates:
    Cell(row, 'A').value = name
    Cell(row, 'B').value = table
    Cell(row, 'C').value = score
    row += 1

