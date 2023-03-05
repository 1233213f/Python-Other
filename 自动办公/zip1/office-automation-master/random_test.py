# Author：刘清九
# Data  ：2020/7/25
"""
生成35份随机试卷，每份试卷五道题
不同试卷的题目顺序不一样
相同题目的选项不一样，正确答案位置不一样
"""
import random

capitals = {'北京': '北京',
            'liaoning': 'shenyang',
            'heilongjiang': 'haerbin',
            'jilin': 'changchun',
            '江苏': '南京',
            '新疆': 'wulumuqi'}
# 生成35份试卷
for quiz_num in range(15):
    # 建立文件
    quiz_file = open('quiz_file%s.txt' % (quiz_num + 1), 'w')
    quiz_file_answer = open('quiz_file_answer%s.txt' % (quiz_num + 1), 'w')
    # 写上标题和头部
    quiz_file.write('Quiz %s\n' % (quiz_num + 1))
    quiz_file_answer.write('Quiz Answer %s\n\n' % (quiz_num + 1))

    quiz_file.write('\nName:\n\nDate:\n\nPeriod:\n\n')

    # 置乱州的顺序
    states = list(capitals.keys())
    random.shuffle(states)

    # 循环5个州，并且为每个州创建问题和答案
    for question_num in range(5):
        correct_answer = capitals[states[question_num]]
        wrong_answers = list(capitals.values())
        del wrong_answers[wrong_answers.index(correct_answer)]
        wrong_answers = random.sample(wrong_answers, 3)
        answer_options = wrong_answers + [correct_answer]
        random.shuffle(answer_options)
        # 将内容写入测验试卷
        quiz_file.write('%s What\'s the capital of %s?\n' % (question_num + 1, states[question_num]))
        for i in range(4):
            quiz_file.write('%s. %s\n' % ('ABCD'[i], answer_options[i]))
        quiz_file.write('\n')
        # 将答案写入文件
        quiz_file_answer.write('%s. %s\n\n' % (question_num + 1, 'ABCD'[answer_options.index(correct_answer)]))

    quiz_file.close()
    quiz_file_answer.close()
