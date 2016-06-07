import pypyodbc
from CONSTANTS import CONNECTION_STRING, SQL_COMMAND


class Node:
    def __init__(self, init_data):
        self.data = init_data
        self.next = None

    def get_data(self):
        return self.data

    def get_next(self):
        return self.next

    def set_data(self, new_data):
        self.data = new_data

    def set_next(self, new_next):
        self.next = new_next


class SingleLinkList:
    def __init__(self):
        self.head = None

    def is_empty(self):
        return self.head is None

    def add_front(self, item):
        temp = Node(item)
        temp.set_next(self.head)
        self.head = temp

    def add_back(self, item):
        last_position = self.head
        if last_position is None:
            temp = Node(item)
            temp.set_next(self.head)
            self.head = temp
        else:
            while last_position.get_next() is not None:
                last_position = last_position.get_next()
            last_position.set_next(Node(item))

    def size(self):
        current = self.head
        count = 0
        while current is not None:
            count += 1
            current = current.get_next()

        return count

    def print(self):
        current = self.head
        while current is not None:
            for data in current.get_data():
                data.__str__()
            current = current.get_next()

    def push_to_sql(self):
        try:
            cnx = pypyodbc.connect(CONNECTION_STRING)
            cur = cnx.cursor()
            current = self.head
            while current is not None:
                for data in current.get_data():
                    values = data.get_sql_string()
                    try:
                        cur.execute(SQL_COMMAND, values)
                        cur.commit()
                    except pypyodbc.IntegrityError:
                        print("Record Already exists: ", values[-1])
                        pass
                current = current.get_next()
        except Exception as e:
            print("General exception pushing to SQL")
            print(e)
        else:
            cur.close()
            cnx.close()

    # Deprecated for client_container
    def search(self, item):
        current = self.head
        found = False
        while current is not None and not found:
            if current.get_data() == item:
                found = True
            else:
                current = current.get_next()

        return found

    def remove(self, item):
        current = self.head
        previous = None
        found = False
        while not found:
            if current.get_data() == item:
                found = True
            else:
                previous = current
                current = current.get_next()

        if previous is None:
            self.head = current.get_next()
        else:
            previous.setNext(current.get_next())
