import random
from typing import List, Dict, Set, Tuple, Optional
from dataclasses import dataclass
import json
from datetime import datetime, timedelta
import itertools
import pandas as pd


@dataclass
class Attendee:
    name: str
    gender: str
    seniority: str
    field: str
    assign_head_table: bool = False

    def __hash__(self):
        return hash(self.name)


@dataclass
class TableConstraints:
    min_seats: int
    max_seats: int

    def __post_init__(self):
        if self.min_seats > self.max_seats:
            raise ValueError("Minimum seats cannot be greater than maximum seats")
        if self.min_seats < 2:
            raise ValueError("Minimum seats must be at least 2")


class SeatingHistory:
    def __init__(self, filename: str = "seating_history.json", memory_events: int = 3):
        self.filename = filename
        self.memory_events = memory_events
        self.history = self.load_history()  # Changed from _load_history to load_history

    def load_history(self) -> List[Dict]:  # Changed from _load_history to load_history
        try:
            with open(self.filename, 'r') as f:
                history = json.load(f)
                # Convert string dates to datetime objects
                for event in history:
                    event['date'] = datetime.fromisoformat(event['date'])
                return history
        except FileNotFoundError:
            return []

    def save_arrangement(self, arrangement: List[List[Attendee]], date: datetime = None):
        if date is None:
            date = datetime.now()

        serializable_arrangement = [
            [{"name": a.name, "gender": a.gender, "seniority": a.seniority,
              "field": a.field, "assign_head_table": a.assign_head_table}
             for a in table]
            for table in arrangement
        ]

        self.history.append({
            "date": date,
            "arrangement": serializable_arrangement
        })

        self.history.sort(key=lambda x: x['date'])

        serializable_history = [
            {**event, 'date': event['date'].isoformat()}
            for event in self.history
        ]
        with open(self.filename, 'w') as f:
            json.dump(serializable_history, f, indent=2)

    def get_recent_pairings(self) -> Dict[Tuple[str, str], List[int]]:
        recent_events = self.history[-self.memory_events:] if self.history else []
        pairings = {}

        all_attendees = set()
        for event in recent_events:
            for table in event["arrangement"]:
                all_attendees.update(person["name"] for person in table)

        for name1, name2 in itertools.combinations(all_attendees, 2):
            pair = tuple(sorted([name1, name2]))
            pairings[pair] = [0] * len(recent_events)

        for event_idx, event in enumerate(recent_events):
            for table in event["arrangement"]:
                names = [person["name"] for person in table]
                for name1, name2 in itertools.combinations(names, 2):
                    pair = tuple(sorted([name1, name2]))
                    pairings[pair][event_idx] = 1

        return pairings

class SeatingOptimizer:
    def __init__(self, constraints: TableConstraints, history: SeatingHistory = None):
        self.constraints = constraints
        self.history = history or SeatingHistory()

    def calculate_table_diversity_score(self, table: List[Attendee]) -> float:
        if not table or len(table) < self.constraints.min_seats:
            return 0.0

        gender_ratio = sum(1 for a in table if a.gender == 'F') / len(table)
        gender_score = 1 - abs(0.5 - gender_ratio)

        senior_ratio = sum(1 for a in table if a.seniority == 'senior') / len(table)
        seniority_score = 1 - abs(0.5 - senior_ratio)

        unique_fields = len(set(a.field for a in table))
        field_score = unique_fields / len(table)

        return (gender_score + seniority_score + field_score) / 3

    def calculate_time_weighted_penalty(self, table: List[Attendee],
                                        recent_pairings: Dict[Tuple[str, str], List[int]]) -> float:
        if not table or len(table) < self.constraints.min_seats:
            return 0.0

        total_penalty = 0
        num_pairs = 0
        time_weights = [1.0, 0.6, 0.3]

        for person1, person2 in itertools.combinations(table, 2):
            pair = tuple(sorted([person1.name, person2.name]))
            if pair in recent_pairings:
                pair_history = recent_pairings[pair]
                weighted_penalty = sum(w * h for w, h in zip(time_weights[:len(pair_history)], pair_history))
                total_penalty += weighted_penalty
                num_pairs += 1

        if num_pairs == 0:
            return 1.0

        avg_penalty = total_penalty / num_pairs
        return max(0, 1 - avg_penalty)

    def create_balanced_tables(self, attendees: List[Attendee], num_tables: int) -> List[List[Attendee]]:
        """Distribute attendees across tables while respecting min/max constraints"""
        if not attendees:
            return []

        # Calculate target size for even distribution
        total_attendees = len(attendees)
        target_size = total_attendees // num_tables

        # Verify we can create valid tables
        if target_size < self.constraints.min_seats:
            raise ValueError(
                f"Too few attendees ({total_attendees}) to create {num_tables} tables with minimum {self.constraints.min_seats} seats")

        # Initialize empty tables
        tables = [[] for _ in range(num_tables)]
        remaining_attendees = attendees.copy()

        # First pass: ensure minimum sizes
        for table in tables:
            for _ in range(self.constraints.min_seats):
                if remaining_attendees:
                    table.append(remaining_attendees.pop())

        # Second pass: distribute remaining attendees
        while remaining_attendees:
            # Find table with fewest attendees that isn't at max capacity
            valid_tables = [t for t in tables if len(t) < self.constraints.max_seats]
            if not valid_tables:
                raise ValueError("Cannot distribute remaining attendees while respecting maximum table size")

            target_table = min(valid_tables, key=len)
            target_table.append(remaining_attendees.pop())

        return tables

    def optimize_seating(self, attendees: List[Attendee], num_tables: int,
                         head_table_assignments: Optional[List[Attendee]] = None,
                         iterations: int = 1000) -> List[List[Attendee]]:
        recent_pairings = self.history.get_recent_pairings()
        best_arrangement = []
        best_score = -1

        # Handle head table assignments
        head_table_attendees = head_table_assignments or [a for a in attendees if a.assign_head_table]
        other_attendees = [a for a in attendees if a not in head_table_attendees]

        # Verify we have enough total attendees
        total_attendees = len(attendees)
        min_required = self.constraints.max_seats + (self.constraints.min_seats * (num_tables - 1))
        if total_attendees < min_required:
            raise ValueError(
                f"Need at least {min_required} attendees: {self.constraints.max_seats} for head table "
                f"and {self.constraints.min_seats} each for {num_tables - 1} other tables"
            )

        # Verify we don't have too many fixed head table assignments
        if len(head_table_attendees) > self.constraints.max_seats:
            raise ValueError(f"Too many head table assignments (maximum is {self.constraints.max_seats})")

        for _ in range(iterations):
            try:
                # Start with head table assignments
                current_head_table = head_table_attendees.copy()
                available_for_head = other_attendees.copy()

                # Fill head table to exactly max_seats
                if len(current_head_table) < self.constraints.max_seats:
                    random.shuffle(available_for_head)
                    seats_needed = self.constraints.max_seats - len(current_head_table)
                    current_head_table.extend(available_for_head[:seats_needed])
                    available_for_head = available_for_head[seats_needed:]

                # Create other tables with remaining attendees
                random.shuffle(available_for_head)
                other_tables = self.create_balanced_tables(available_for_head, num_tables - 1)

                current_arrangement = [current_head_table] + other_tables

                # Score the arrangement
                diversity_scores = [self.calculate_table_diversity_score(table) for table in current_arrangement]
                recency_scores = [self.calculate_time_weighted_penalty(table, recent_pairings) for table in
                                  current_arrangement]

                # Add size balance score (only for non-head tables)
                other_table_sizes = [len(table) for table in current_arrangement[1:]]
                size_variations = [abs(s1 - s2) for s1, s2 in itertools.combinations(other_table_sizes, 2)]
                size_balance_score = 1.0 - (sum(size_variations) / (
                            len(current_arrangement) * self.constraints.max_seats)) if size_variations else 1.0

                # Combine scores (60% diversity, 25% recency, 15% size balance)
                current_score = (
                        0.6 * sum(diversity_scores) / len(current_arrangement) +
                        0.25 * sum(recency_scores) / len(current_arrangement) +
                        0.15 * size_balance_score
                )

                if current_score > best_score:
                    best_score = current_score
                    best_arrangement = current_arrangement

            except ValueError:
                continue

        if not best_arrangement:
            raise ValueError("Could not find valid seating arrangement")

        # Verify head table size
        if len(best_arrangement[0]) != self.constraints.max_seats:
            raise ValueError("Failed to create head table with required size")

        return best_arrangement


def print_seating_arrangement(arrangement: List[List[Attendee]], history: SeatingHistory = None):
    recent_pairings = history.get_recent_pairings() if history else None

    # Print head table
    print("\nHead Table (Fixed Size):")
    print("-" * 50)
    print(f"Number of guests: {len(arrangement[0])} (Maximum capacity)")
    for attendee in arrangement[0]:
        assigned_marker = "*" if attendee.assign_head_table else " "
        print(f"{attendee.name:15} | {attendee.gender:6} | {attendee.seniority:6} | {attendee.field} {assigned_marker}")

    # Print other tables
    for i, table in enumerate(arrangement[1:], 2):
        print(f"\nTable {i}:")
        print("-" * 50)
        print(
            f"Number of guests: {len(table)} (Flexible capacity)")
        for attendee in table:
            print(f"{attendee.name:15} | {attendee.gender:6} | {attendee.seniority:6} | {attendee.field}")

        if recent_pairings:
            print("\nRecent interactions at this table:")
            for person1, person2 in itertools.combinations(table, 2):
                pair = tuple(sorted([person1.name, person2.name]))
                if pair in recent_pairings:
                    history = recent_pairings[pair]
                    recent_count = sum(history)
                    if recent_count > 0:
                        events_ago = [i + 1 for i, sat in enumerate(reversed(history)) if sat]
                        events_text = ", ".join(f"{n} event{'s' if n > 1 else ''} ago" for n in events_ago)
                        print(f"{person1.name} and {person2.name} sat together: {events_text}")


def Attendees_from_spreadsheet(excelfile):
    try:
        df = pd.read_excel(excelfile)
        df_attending = df[df.attending == 'Y']
        attendees = []
        for ind, row in df_attending.iterrows():
            head_table = True if row['head table']==1.0 else False
            attendees.append(Attendee(row['name'], row['gender'], row['seniority'],
                                      row['division'], assign_head_table=head_table))
        return attendees
    except FileNotFoundError:
        return []
def main():
    # Initialize with table size constraints
    TABLESIZE = 6
    constraints = TableConstraints(min_seats=4, max_seats=TABLESIZE)
    history = SeatingHistory(memory_events=3)
    optimizer = SeatingOptimizer(constraints=constraints, history=history)

    # Create sample attendees
 #   attendees = [
 #       Attendee("Dr. Smith", "F", "senior", "Physics", assign_head_table=True),
 #       Attendee("Dr. Johnson", "M", "senior", "Biology", assign_head_table=True),
 #       Attendee("Ms. Chen", "F", "junior", "Computer Science"),
 #       Attendee("Mr. Brown", "M", "junior", "Physics"),
 #       Attendee("Dr. Garcia", "F", "senior", "Chemistry"),
 #       Attendee("Mr. Wilson", "M", "junior", "Biology"),
 #       Attendee("Dr. Lee", "F", "senior", "Mathematics"),
 #       Attendee("Ms. Taylor", "F", "junior", "Chemistry"),
 #       Attendee("Dr. Anderson", "M", "senior", "Computer Science"),
 #       Attendee("Mr. Martinez", "M", "junior", "Mathematics"),
 #       Attendee("Dr. White", "F", "senior", "Physics"),
 #       Attendee("Ms. Lopez", "F", "junior", "Biology")
 #   ]

    attendees = Attendees_from_spreadsheet('MSF_2024-25.xlsx')

    # Generate optimal seating arrangement
    Nguests = len(attendees)
    Ntables = int(Nguests/TABLESIZE + 0.5)
    arrangement = optimizer.optimize_seating(attendees, num_tables=Ntables+1)

    # Print the results
    print("\nOptimal Seating Arrangement:")
    print_seating_arrangement(arrangement, history)

    # Save this arrangement to history
    history.save_arrangement(arrangement)


if __name__ == "__main__":
    main()