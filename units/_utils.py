def line_counter(text, capacity_per_line):
    lines = text.split("\n")

    count = 0
    for line in lines:
        length = 0
        for word in line.split():
            length += len(word)
            if length > capacity_per_line:
                if len(word) <= capacity_per_line:
                    count += 1
                    length = len(word)
                else:
                    # encounter a long word(> capacity)
                    count += word // capacity_per_line
                    length = word % capacity_per_line

        if length != 0:
            count += 1

    return count

if __name__ == "__main__":
    pass
