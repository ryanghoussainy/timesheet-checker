from datetime import date

class Entry:
    def __init__(self, date: date, hours: float, rate: float):
        self.date = date
        self.hours = hours
        self.rate = rate
    
    def __eq__(self, other):
        if not isinstance(other, Entry):
            return NotImplemented
        return (self.date, self.hours, self.rate) == (other.date, other.hours, other.rate)

    def __hash__(self):
        return hash((self.date, self.hours, self.rate))
    
    def __iter__(self):
        return iter((self.date, self.hours, self.rate))
