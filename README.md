# ENG

## excel-vba-ngram
Excel VBA user defined functions for N-GRAM similarity text analysis. There are two types of functions base on two types of  similaryty metrics: fast and good. 
**Download and open N-GRAMS.xlsm** for discription and examples.

N-GRAMS.xlam - is identical to N-GRAMS.xlsm but ready for plug in as add-in.

## Good
But slow. Builds vector space for each comparison. Builds vectros is this space for two compared string. Uses cosine similarity between vectors.

Functions:
 - =VLOOKUP_NGRAM(...) - mimics regular VLOOKUP (but returns value only from searching range - 1st column)
 - =VLOOKUP_NGRAM_help() -  use this for detailed help inside Excel
 - TextSimilarity(...) - similarity between two strings

## Fast
The method above is better but could be slow. This functions are simplier and faster: for text similarity they  count the number of equal NGRAMs between string.

Functions:
 - =VLOOKUP_NGRAM_fast(...) - mimics regular VLOOKUP (but returns value only from searching range - 1st column)
 - =VLOOKUP_NGRAM_fast_help() -  use this for detailed help inside Excel
 - TextSimilarityFast(...) - similarity between two strings

# RUS

## excel-vba-ngram
Функции Excel VBA для анализа текста и его схожести с помощью NGRAM. Два вида фукнций: хорошие и быстрые. 
**Скачайте и откройте N-GRAMS.xlsm** для детального описания и примеров.

N-GRAMS.xlam - тоже самое что и N-GRAMS.xlsm но сохранено как в формате надстройки Excel - готово для подключения как надстройка.

## Хорошие
Но медленные. Строят векторное пространство NGRAM для каждого стравнения двух строк. Строят по вектору в этом пространстве для каждой сравниваемой строки. Использует cosine similarity как метрику близостри векторов/строк.

Функции:
 - =ВПР_NGRAM(...) - работает почти как обычный ВПР, но выдает значение не из произвольного столбца а только из первого - в котором и ищет.
 - =ВПР_NGRAM_помощь() -  наберите это для получения справки по функции в Экселе.
 - TextSimilarity(...) - вычисляет схожесть двух строк.

## Быстрые
Предыдущие методы сравнения хороши, но могут быть оченб медленными. Эти функции проще: они просто считают количество одинаковых NGRAM в сравниваемых строках и спользют это поличество как метрику близости строк.

Функции:
 - =ВПР_NGRAM_быстрый(...) - работает почти как обычный ВПР, но выдает значение не из произвольного столбца а только из первого - в котором и ищет.
 - =ВПР_NGRAM_быстрый_помощь() -  наберите это для получения справки по функции в Экселе.
 - TextSimilarityFast(...) - вычисляет схожесть двух строк.
