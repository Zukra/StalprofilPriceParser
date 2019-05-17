<?php


namespace App;


use PHPExcel;

class PriceParser {

    /**
     * RIBBED BARS
     * @param PHPExcel $objPHPExcel
     * @return array
     */
    public static function getRibberBars(PHPExcel $objPHPExcel) {
        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('RIBBED BARS & WIRE ROD');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[0])) == "RIBBED BARS") {
                // ставим флаг в true
                $startData = true;
                // переходим к следующей итерации, где уже непосредственно будет работа с данными
                continue;
            }

            if ($startData && ! trim($rowData[0])) { // цены по этой категории закончились
                break;
            }

            // если данные еще не начались либо попалась пустая строка
            if (! $startData || ! trim($rowData[0])) {
                // переходим к следующей итерации
                continue;
            }

            // убираем лишние пробелы, заменяем запятые на точки
            $s = str_replace([' ', ','], ['', '.'], $rowData[0]);

            // получаем массив диаметров
            if (strpos($s, '-') !== false) {
                list($from, $to) = explode('-', $s);
                $diameters = range($from, $to);
            } elseif (strpos($s, ';') !== false) {
                $diameters = explode(';', $s);
            } else {
                $diameters = [$s];
            }

            // цена
            $steel = trim($rowData[1]);

            // создаем специальный хеш для хранения цен
            foreach ($diameters as $d) {
                $x[$d] = $steel;
            }
        }

        return $x;
    }

    /**
     * WIRE ROD
     * @param PHPExcel $objPHPExcel
     * @return array
     */
    public static function getWireRod(PHPExcel $objPHPExcel) {
        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('RIBBED BARS & WIRE ROD');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[0])) == 'WIRE ROD') {
                // ставим флаг в true
                $startData = true;
                // переходим к следующей итерации, где уже непосредственно будет работа с данными
                continue;
            }

            if ($startData && ! trim($rowData[0])) { // цены по этой категории закончились
                break;
            }

            // если данные еще не начались либо попалась пустая строка
            if (! $startData || ! trim($rowData[0])) {
                // переходим к следующей итерации
                continue;
            }

            // убираем лишние пробелы, заменяем запятые на точки
            $s = str_replace([' ', ','], ['', '.'], $rowData[0]);

            // получаем массив диаметров
            if (strpos($s, '-') !== false) {
                list($from, $to) = explode('-', $s);
                $diameters = range($from, $to, 0.5);
            } elseif (strpos($s, ';') !== false) {
                $diameters = explode(';', $s);
            } else {
                $diameters = [$s];
            }

            // цена
            $steel = trim($rowData[1]);

            // создаем специальный хеш для хранения цен
            foreach ($diameters as $d) {
                $x[$d . ''] = $steel;
            }
        }

        return $x;
    }

    /**
     * SHEETS
     * @param PHPExcel $objPHPExcel
     * @return array
     */
    public static function getSheets(PHPExcel $objPHPExcel) {
        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('SHEETS & PLATES');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[0])) == 'SHEETS') {
                // ставим флаг в true
                $startData = true;
                // переходим к следующей итерации, где уже непосредственно будет работа с данными
                continue;
            }

            if ($startData && ! trim($rowData[0])) { // цены по этой категории закончились
                break;
            }

            // если данные еще не начались либо попалась пустая строка или заголовок таблицы
            if (! $startData || ! trim($rowData[0]) || trim($rowData[0]) == 'THICKNESS  [mm]') {
                // переходим к следующей итерации
                continue;
            }

            // получаем массив толщин строки
            if (strpos($rowData[0], '÷') !== false) { // это диапазон
                $range = explode('÷', $rowData[0]);
                $thicknesses = range((float)$range[0], (float)$range[1]);
            } else { // это просто число
                $thicknesses = [(float)trim($rowData[0])];
            }

            // получаем массив размеров
            if (strpos($rowData[1], 'size acc. To production program') !== false) { // универсальный размер
                $sizes = ['*'];
            } else { // перечисление размеров
                // убираем лишние пробелы
                $sizes = trim(preg_replace('!\s+!', ' ', $rowData[1]));
                // убираем оплошности заполнявшего прайс
                $sizes = str_replace('2500 5000', '2500x5000', $sizes);
                // режем строку в массив
                $sizes = explode(' ', $sizes);
            }

            // цена за сталь S235JR, S235JRC
            $steel1 = trim($rowData[2]);
            // цена за сталь S355J2, ST52-3, S355J2C
            $steel2 = trim($rowData[3]);

            // создаем специальный хеш для хранения цен
            foreach ($thicknesses as $t) {
                foreach ($sizes as $s) {
                    $x["{$t}x{$s}xS235JR"] = $steel1;
                    $x["{$t}x{$s}xS235JRC"] = $steel1;
                    $x["{$t}x{$s}xS355J2"] = $steel2;
                    $x["{$t}x{$s}xS355J2C"] = $steel2;
                    $x["{$t}x{$s}xST52-3"] = $steel2;
                    $x["{$t}x{$s}xS355JR"] = $steel2;
                }
            }
        }

        return $x;
    }

    public static function getColdRolledPlates(PHPExcel $objPHPExcel) {
        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('SHEETS & PLATES');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[0])) == 'COLD ROLLED PLATES') {
                // ставим флаг в true
                $startData = true;
                // переходим к следующей итерации, где уже непосредственно будет работа с данными
                continue;
            }
            if ($startData && ! trim($rowData[0])) { // цены по этой категории закончились
                break;
            }
            // если данные еще не начались либо попалась пустая строка или заголовок таблицы
            if (! $startData || ! trim($rowData[0]) || trim($rowData[0]) == 'THICKNESS  [mm]') {
                // переходим к следующей итерации
                continue;
            }
            // получаем массив толщин строки
            if (strpos($rowData[0], '÷') !== false) { // это диапазон
                $range = explode('÷', $rowData[0]);
                $thicknesses = range((float)$range[0], (float)$range[1], 0.5);
            } else { // это просто число
                $thicknesses = [(float)trim($rowData[0])];
            }
            // получаем массив размеров
            // убираем лишние пробелы
            $sizes = trim(preg_replace('!\s+!', ' ', $rowData[1]));
            // режем строку в массив
            $sizes = explode(' ', $sizes);

            // цена за сталь DC01
            $steel = trim($rowData[2]);

            // создаем специальный хеш для хранения цен
            foreach ($thicknesses as $t) {
                foreach ($sizes as $s) {
                    $x["{$t}x{$s}xDC01"] = $steel;
                }
            }
        }

        return $x;
    }

    /**
     * UNEQUAL ANGLES
     * @param PHPExcel $objPHPExcel
     * @return array
     */
    public static function getUnequalAngles(PHPExcel $objPHPExcel) {
        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('Merchant Bars');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[1])) == 'UNEQUAL ANGLES') {
                // ставим флаг в true
                $startData = true;
                // переходим к следующей итерации, где уже непосредственно будет работа с данными
                continue;
            }

            if ($startData && ! trim($rowData[1])) { // цены по этой категории закончились
                break;
            }

            // если данные еще не начались либо попалась пустая строка
            if (! $startData || ! trim($rowData[1])) {
                // переходим к следующей итерации
                continue;
            }

            // размеры
            // заменяем фразу "oraz" и запятые на точку с запятой
            // избавляемся от пробелов, приводим к верхнему регистру
            $sizes = strtoupper(str_replace([' ', 'oraz', ','], ['', ';', ';'], $rowData[1]));
            // убираем косяки и частные случаи
            $sizes = str_replace(['160X80150X75', 'L=6MB'], ['160X80;150X75;', 'X6000'], $sizes);
            // разбиваем в массив
            $sizes = explode(';', $sizes);
            // пост-обход массива
            foreach ($sizes as $k => $s) {
                // в ячейке есть дефис, нужно разобрать диапазон
                if (strpos($s, '-') !== false) {
                    // разбиваем строку на две части
                    $a = substr($s, 0, strrpos($s, 'X'));
                    $b = substr($s, strrpos($s, 'X') + 1, strlen($s) - strrpos($s, 'X'));
                    // во второй части диапазон, получаем его
                    list($p, $q) = explode('-', $b);
                    // добавляем в массив размеры по диапазону
                    for ($i = $p; $i <= $q; $i++) {
                        $sizes[] = $a . 'X' . $i;
                    }
                    // удаляем ненужный элемент
                    unset($sizes[$k]);
                }
                // в ячейке есть амперсанд, нужно сделать из одного размера два
                if (strpos($s, '&') !== false) {
                    // разбиваем строку на две части
                    list($a, $b) = explode('&', $s);
                    // левая часть - самодостаточный размер
                    $sizes[] = $a;
                    // другой получаем из части левого без последнего фрагмента + правая часть
                    $sizes[] = substr($a, 0, strrpos($a, 'X') + 1) . $b;
                    // удаляем ненужный элемент
                    unset($sizes[$k]);
                }
            }

            // цена за сталь S235JR, S235JR/S275JR
            $steel1 = trim($rowData[2]);
            // цена за сталь S355J2
            $steel2 = trim($rowData[3]);

            // создаем специальный хеш для хранения цен
            foreach ($sizes as $s) {
                $x["{$s}xS235JR"] = $steel1;
                $x["{$s}xS275JR"] = $steel1;
                $x["{$s}xS235JR/S275JR"] = $steel1;
                $x["{$s}xS355J2"] = $steel2;
            }
        }

        return $x;
    }

    /**
     * EQUAL-LEG ANGLES
     * @param PHPExcel $objPHPExcel
     * @return array
     */
    public static function getEqualLegAngles(PHPExcel $objPHPExcel) {
        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('Merchant Bars');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[1])) == 'EQUAL-LEG ANGLES') {
                // ставим флаг в true
                $startData = true;
                // переходим к следующей итерации, где уже непосредственно будет работа с данными
                continue;
            }

            if ($startData && ! trim($rowData[1])) { // цены по этой категории закончились
                break;
            }

            // если данные еще не начались либо попалась пустая строка
            if (! $startData || ! trim($rowData[1])) {
                // переходим к следующей итерации
                continue;
            }

            // размеры
            // заменяем фразу "oraz" и запятые на точку с запятой
            // избавляемся от пробелов, приводим к верхнему регистру
            if (strpos($rowData[1], 'x') === false) {
                $sizes = strtoupper(str_replace([' ', 'oraz', ','], ['', ';', ';'], $rowData[1]));
            } else {
                $sizes = strtoupper(str_replace([' ', 'oraz', ','], ['', ';', '&'], $rowData[1]));
            }
            if (strpos($sizes, 'X') !== false && strpos($sizes, ';') !== false && strpos($sizes, '-') !== false) {
                $tmp = substr($sizes, strpos($sizes, 'X'));
                $sizes = str_replace([';'], [$tmp . ';'], $sizes);
                unset($tmp);
            }

            // разбиваем в массив
            $sizes = explode(';', $sizes);
            // пост-обход массива
            foreach ($sizes as $k => $s) {
                // в ячейке есть дефис, нужно разобрать диапазон
                if (strpos($s, '-') !== false) {
                    // разбиваем строку на две части
                    $a = substr($s, 0, strrpos($s, 'X'));
                    $b = substr($s, strrpos($s, 'X') + 1, strlen($s) - strrpos($s, 'X'));
                    // во второй части диапазон, получаем его
                    list($p, $q) = explode('-', $b);
                    // добавляем в массив размеры по диапазону
                    for ($i = $p; $i <= $q; $i++) {
                        $sizes[] = $a . 'X' . $i;
                    }
                    // удаляем ненужный элемент
                    unset($sizes[$k]);
                }
                // в ячейке есть амперсанд, нужно сделать из одного размера два
                if (strpos($s, '&') !== false) {
                    $tmp = explode('&', $s);
                    foreach ($tmp as $key => $item) {
                        $sizes[] = ($key == 0)
                            ? $item
                            : substr($tmp[0], 0, strrpos($tmp[0], 'X') + 1) . $item;
                    }

                    /*// разбиваем строку на две части
                    list($a, $b) = explode('&', $s);
                    // левая часть - самодостаточный размер
                    $sizes[] = $a;
                    // другой получаем из части левого без последнего фрагмента + правая часть
                    $sizes[] = substr($a, 0, strrpos($a, 'X') + 1) . $b;
                    // удаляем ненужный элемент*/

                    unset($sizes[$k]);
                }
            }

            // цена за сталь S235JR, S235JR/S275JR
            $steel1 = trim($rowData[2]);
            // цена за сталь S355J2
            $steel2 = trim($rowData[3]);

            // создаем специальный хеш для хранения цен
            foreach ($sizes as $s) {
                $x["{$s}xS235JR"] = $steel1;
                $x["{$s}xS235JR/S275JR"] = $steel1;
                $x["{$s}xS355J2"] = $steel2;
            }
        }

        return $x;
    }

    /**
     * Cold formed Hollow sections EN 10219
     * @param PHPExcel $objPHPExcel
     * @return array
     */
    public static function getColdFormedHollowSections(PHPExcel $objPHPExcel) {
        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('Hollow Sections ');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (trim($rowData[1]) == 'Cold formed Hollow sections EN 10219') {
                // ставим флаг в true
                $startData = true;
                // переходим к следующей итерации, где уже непосредственно будет работа с данными
                continue;
            }

            if ($startData && ! trim($rowData[1])) { // цены по этой категории закончились
                break;
            }

            // если данные еще не начались либо попалась пустая строка
            if (! $startData || ! trim($rowData[1])) {
                // переходим к следующей итерации
                continue;
            }

            // разбиваем строку на две части
            list($a, $b) = explode('thickness', $rowData[1]);

            // в левой части диапазон ширин
            preg_match('/(from (\d+))/', $a, $matches);
            // левая граница диапазона
            if ($matches[2]) {
                $from = (int)$matches[2];
            } else {
//                preg_match('/(\d+)/', $a, $matches);
                $from = 1;
            }

            // правая граница
            preg_match('/(to (\d+))/', $a, $matches);
            if ($matches[2]) {
                $to = (int)$matches[2];
            } else {
                $to = 1000;
            }

            preg_match('/(over (\d+))/', $a, $matches);
            // правая граница
            if ($matches[2]) {
                $from = $matches[2] + 1;
            }

            preg_match('/to|from|over/', $a, $matches);
            if (empty($matches) && strpos($a, 'x') === false) {
                preg_match('/\d+/', $a, $matches);
                if (! empty($matches[0])) {
                    $from = $matches[0];
                    $to = $matches[0];
                }
            }

            // массив ширин
//            $widthes = range($from, $to + 1);
            $widthes = range($from, $to);

            // в правой перечень толщин
            // избавляемся от лишних символов
            $b = str_replace(['mm', ' ', ','], ['', '', '.'], $b);
            // получаем массив толщин
            $thicknesses = explode('&', $b);
            // преобразуем к вещественному числу
            foreach ($thicknesses as $k => $t) {
                $thicknesses[$k] = (float)$t;
            }

            // цена за сталь S235JR(H), E235
            $steel1 = trim($rowData[2]);
            // цена за сталь S355J2(H)
            $steel2 = trim($rowData[3]);

            // создаем специальный хеш для хранения цен
            foreach ($widthes as $w) {
                foreach ($thicknesses as $t) {
                    $x["{$w}x{$t}xEN10219xS235JR"] = $steel1;
                    $x["{$w}x{$t}xEN10219xS235JRH"] = $steel1;
                    $x["{$w}x{$t}xEN10219xE235"] = $steel1;
                    $x["{$w}x{$t}xEN10219xS355J2"] = $steel2;
                    $x["{$w}x{$t}xEN10219xS355J2H"] = $steel2;
                    $x["{$w}x{$t}xEN10219xS355J2H/S420MH"] = $steel2;
                }
            }
        }

        return $x;
    }
}