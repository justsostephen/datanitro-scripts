"""
Calculate combined Say and VP delegate scores.

"""

def say_scores(sheet, correct):
    """Calculate and compile list of Say scores."""
    active_sheet(sheet)
    delegates = []
    # Set data start row.
    row = 2
    # Loop through rows until empty row is found.
    data_present = True
    while data_present:
        # Append nested list to 'delegates' list containing delegate PIN,
        # answers and initialised score.
        if sheet in ('asthma', 'copd', 'ipf'):
            delegates.append(
                [Cell(row, 'C').value,
                CellRange((row, 'D'), (row, 'F')).value, 0])
        elif sheet == 'walkabout_1':
            delegates.append(
                [int(Cell(row, 'A').value[1:]),
                CellRange((row, 'C'), (row, 'E')).value, 0])
        elif sheet == 'walkabout_2':
            delegates.append(
                [int(Cell(row, 'A').value[1:]), [Cell(row, 'C').value], 0])
        else:
            delegates.append(
                [int(Cell(row, 'A').value[1:]), [Cell(row, 'B').value], 0])
        # Compare delegate answers to correct answers, add 10 to score for each
        # match.
        i = 0
        for answer in correct:
            if answer == delegates[-1][1][i]:
                delegates[-1][2] += 10
            i += 1
        # Write delegate score to sheet.
        Cell(row, 'H').value = delegates[-1][2]
        row += 1
        # Check next row for data.
        data_present = Cell(row, 'A').value
    # Add column header.
    Cell(1, 'H').value = 'Score'
    return delegates

def vp_scores(sheet):
    """Compile list of VP scores."""
    active_sheet(sheet)
    delegates = []
    # Set data start row.
    row = 6
    # Loop through rows until empty row is found.
    data_present = True
    while data_present:
        # Append nested list to 'delegates' list containing delegate PIN and
        # score.
        delegates.append(CellRange((row, 'H'), (row, 'I')).value)
        row += 1
        # Check next row for data.
        data_present = Cell(row, 'A').value
    return delegates

def scoreboard():
    """Calculate and compile combined delegate scores."""
    active_sheet('scores')
    # Read delegate list from sheet.
    delegates = CellRange((2, 'A'), (192, 'A')).value
    totals = []
    i = 0
    print 'BEGUN'
    # Add delegate scores from each sheet.
    for delegate in delegates:
        # Initialise delegate total score.
        total = 0
        # If there's data on the sheet, add sheet score to total. Say correct
        # answers given in second argument to 'say_scores'.
        if Cell('asthma', 1, "A").value:
            total = add_say(say_scores('asthma', [3, 1, 2]), delegate, total)
        if Cell('copd', 1, "A").value:
            total = add_say(say_scores('copd', [4, 1, 5]), delegate, total)
        if Cell('ipf', 1, "A").value:
            total = add_say(say_scores('ipf', [1, 4, 2]), delegate, total)
        if Cell('walkabout_1', 1, "A").value:
            total = add_say(say_scores('walkabout_1', [2, 2, 3]), delegate, total)
        if Cell('walkabout_2', 1, "A").value:
            total = add_say(say_scores('walkabout_2', [1]), delegate, total)
        if Cell('walkabout_4', 1, "A").value:
            total = add_say(say_scores('walkabout_4', [3]), delegate, total)
        if Cell('walkabout_7', 1, "A").value:
            total = add_say(say_scores('walkabout_7', [2]), delegate, total)
        if Cell('walkabout_12', 1, "A").value:
            total = add_say(say_scores('walkabout_12', [3]), delegate, total)
        if Cell('walkabout_13', 1, "A").value:
            total = add_say(say_scores('walkabout_13', [1]), delegate, total)
        if Cell('vp_1', 1, "A").value:
            total = add_vp(vp_scores('vp_1'), delegate, total)
        if Cell('vp_2', 1, "A").value:
            total = add_vp(vp_scores('vp_2'), delegate, total)
        if Cell('vp_3', 1, "A").value:
            total = add_vp(vp_scores('vp_3'), delegate, total)
        # Add delegate's total score to 'totals' list.
        totals.append(total)
        i += 1
        print i, 'PROCESSED'
    # Write 'totals' list to sheet.
    CellRange('scores', (2, 'E'), (192, 'E')).value = totals

def add_say(project, delegate, total):
    """Add delegate Say scores to total."""
    for pin, answers, score in project:
        if delegate == pin:
            total += score
            break
    return total

def add_vp(day, delegate, total):
    """Add delegate VP scores to total."""
    for pin, score in day:
        if pin:
            if delegate == int(pin):    # VP outputs 'pin' as a string.
                total += score
                break
    return total

scoreboard()

raw_input('COMPLETE')

