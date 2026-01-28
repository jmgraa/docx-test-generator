from random import shuffle


def shuffle_answers(question):
    order = [i for i in range(len(question.answers))]
    shuffle(order)
    shuffle_answers = [question.answers[o] for o in order]

    return (shuffle_answers, order)
