from django.http import HttpResponse, Http404, HttpResponseRedirect
from django.shortcuts import render, get_object_or_404

from django.urls import reverse

from .models import Choice, Question

from polls.models import Question


def index(request):
    latest_question_list = Question.objects.order_by('-pub_date')[:5]
    context = {'latest_question_list': latest_question_list}
    return render(request, 'polls/index.html', context)

def detail(request, question_id):
    question = get_object_or_404(Question, pk=question_id)
    return render(request, 'polls/detail.html', {'question': question})

def results(request, question_id):
    question = get_object_or_404(Question, pk=question_id)
    return render(request, 'polls/results.html', {'question': question})

def vote(request, question_id):
    question = get_object_or_404(Question, pk=question_id)
    try:
        selected_choice = question.choice_set.get(pk=request.POST['choice'])
    except (KeyError, Choice.DoesNotExist):
        # Redisplay the question voting form.
        return render(request, 'polls/detail.html', {
            'question': question,
            'error_message': "You didn't select a choice.",
        })
    else:
        selected_choice.votes += 1
        selected_choice.save()
        # Always return an HttpResponseRedirect after successfully dealing
        # with POST data. This prevents data from being posted twice if a
        # user hits the Back button.
        return HttpResponseRedirect(reverse('polls:results', args=(question.id,)))

def column_index_value(cname, cols):
    index = 0
    for col in cols:
        if col.value == cname:
            return index
        index += 1
    return -1

def import_excel(request):
    if request.method == "POST":
        file = request.FILES['file2']
        file2 = request.FILES['file']
        stops1 = xlrd.open_workbook(file_contents=file.read())
        stops2 = xlrd.open_workbook(file_contents=file2.read())
        sh = stops1.sheet_by_index(0)
        col_header = sh.row(0)
        tc_index = column_index_value('Öğrenci TC', col_header)
        date_index = column_index_value('Doğum Tarihi', col_header)
        for sheet in stops1.sheets():
            for index in range(0, sheet.nrows):
                try:
                    if index == 0:
                        continue
                    row = sheet.row(index)

                    tc = str(row[tc_index].value).strip()
                    date = str(row[date_index].value).strip()
                    sh2 = stops2.sheet_by_index(0)
                    col_header2 = sh2.row(0)
                    tc_index2 = column_index_value('Öğrenci TC', col_header2)
                    date_index2 = column_index_value('Doğum Tarihi', col_header2)
                    output = 'output.csv'
                    op = open(output, 'wb')
                    wr = csv.writer(op, quoting=csv.QUOTE_ALL)

                    for sheet2 in stops2.sheets():
                        for index2 in range(0, sheet.nrows):
                            try:
                                if index2 == 0:
                                    continue
                                row2 = sheet2.row(index2)
                                tc2 = str(row2[tc_index2].value).strip()
                                date2 = str(row2[date_index2].value).strip()
                                if tc2 == tc:
                                    wr.writerow(date2)

                            except Exception as ex:
                                print(ex)
                except Exception as ex:
                    print(ex)

    return render(request, 'email.html', {'title': 'mail sent'})