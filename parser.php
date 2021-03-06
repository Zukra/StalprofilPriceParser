<?php

require 'vendor/autoload.php';

//file_put_contents('log.log', var_export($matches, true) . "\n", FILE_APPEND);

error_reporting(E_ALL & ~E_NOTICE);

session_start();

// установка кодировки и майм-типа для вывода
header('Content-Type: text/html; charset=utf-8');

// установка неограниченного времени выполнения скрипта
ini_set('max_execution_time', 0);
// ручное задание ограничения потребляемой памяти
ini_set('memory_limit', '256M');

global $prices;

/**
 * конфигурация парсера для конкретного файла
 * содержит в себе настройки, используемые непосредственно при разборе
 */
$parserConfiguration = (object)[
    'namesMap' => (object)[ // карта соответствий наименований
                            'BLACHA OCYNKOWANA'                 => (object)[
                                'ru' => 'Лист оцинкованный',
                                'en' => 'HDG sheet',
                            ],
                            'BLACHA OCYNKOWANA W KRĘGACH'       => (object)[
                                'ru' => 'Рулон оцинкованный',
                                'en' => 'HDG coil',
                            ],
                            'BLACHA G/W'                        => (object)[
                                'ru' => 'Лист г/к',
                                'en' => 'Hot rolled steel plate',
                            ],
                            'BLACHA KOTŁOWA'                    => (object)[
                                'ru' => 'Котельный лист',
                                'en' => 'Boiler plates',
                            ],
                            'BLACHA ŁEZKOWA'                    => (object)[
                                'ru' => 'Рифленый лист',
                                'en' => 'Teardrop hot rolled steel sheet',
                            ],
                            'BLACHA Z/W'                        => (object)[
                                'ru' => 'Лист х/к',
                                'en' => 'Cold rolled steel plate',
                            ],
                            'CEOWNIK EKON. O PÓŁKACH RÓWNOLEG.' => (object)[
                                'ru' => 'Швеллер',
                                'en' => 'Economical chanel with parallel flanges',
                            ],
                            'CEOWNIK PÓŁZAMKNIĘTY'              => (object)[
                                'ru' => 'Коробчатый профиль',
                                'en' => 'Semi-hollow channel section',
                            ],
                            'CEOWNIK Z/G'                       => (object)[
                                'ru' => 'Швеллер гнутый',
                                'en' => 'Cold-formed channel',
                            ],
                            'CEOWNIK'                           => (object)[
                                'ru' => 'Швеллер',
                                'en' => 'Channel',
                            ],
                            'DWUTEOWNIK'                        => (object)[
                                'ru' => 'Балка',
                                'en' => 'Beam',
                            ],
                            'KĄTOWNIK Z/G'                      => (object)[
                                'ru' => 'Уголок гнутый',
                                'en' => 'Cold-formed angle',
                            ],
                            'KĄTOWNIK'                          => (object)[
                                'ru' => 'Уголок',
                                'en' => 'Angle',
                            ],
                            'KSZTAŁTOWNIK KWADRATOWY'           => (object)[
                                'ru' => 'Профиль квадратный',
                                'en' => 'Square hollow section',
                            ],
                            'KSZTAŁTOWNIK PROSTOKĄTNY'          => (object)[
                                'ru' => 'Профиль прямоугольный',
                                'en' => 'Rectangular hollow section',
                            ],
                            'PŁASKOWNIK'                        => (object)[
                                'ru' => 'Полоса',
                                'en' => 'Flat bar',
                            ],
                            'PRĘT GŁADKI CIĄGNIONY'             => (object)[
                                'ru' => 'Круг гладкий тянутый',
                                'en' => 'Round drawn bar',
                            ],
                            'PRĘT GŁADKI'                       => (object)[
                                'ru' => 'Пруток круглый',
                                'en' => 'Round rolled bar',
                            ],
                            'PRĘT KWADRATOWY'                   => (object)[
                                'ru' => 'Пруток квадратный',
                                'en' => 'Square bar',
                            ],
                            'PRĘT ŻEBROWANY'                    => (object)[
                                'ru' => 'Стержень ребристый',
                                'en' => 'Ribbed bars',
                            ],
                            'RURA'                              => (object)[
                                'ru' => 'Труба',
                                'en' => 'Pipe',
                            ],
                            'TEOWNIK'                           => (object)[
                                'ru' => 'Тавр',
                                'en' => 'T-bar',
                            ],
                            'WALCÓWKA'                          => (object)[
                                'ru' => 'Катанка',
                                'en' => 'Rod',
                            ],
    ],
    'prices'   => (object)[ // настройки файла цен
                            'pricesStart' => 9, // стартовая строка цен листа "prices"
    ],
    'titles'   => [
        'name'     => 0,
        'quantity' => 1,
        'unit'     => 2
    ],
];

processForm($parserConfiguration);

/**
 * функция для получения наименования товара на конкретном языке
 *
 * принимает на вход наименование на польском, целевой язык и карту соответствий
 * выбирает нужное значение и возвращает, если оно есть в карте
 *
 * string    namePL    наименование на польском
 * string    lang            язык, на котором нужно получить наименование
 * object    map            карта соответствий
 *
 * return    string        наименование товара на целевом языке
 */
function getName($namePL, $lang, $map) {
    if (isset($map->$namePL->$lang)) {
        return $map->$namePL->$lang;
    } else {
        return '';
    }
}

/**
 * вспомогательная функция для разбора примечания и типа стали
 *
 * принимает на вход строку с содержимым для разбора и объект элемента
 * по ссылке. производит разбор и дописывает нужные свойства объекту
 * функция нужна для сокращения количества кода в функции разбора
 *
 * string    str        строка текста
 * string    item        объект элемента
 *
 * return    void
 */
function parseAdditional($str, &$item) {
    $comment = '';
    // убираем лишние пробелы
    $str = trim($str);

    // в строке есть пробел, следовательно после стали что-то записано
    if (strpos($str, ' ') !== false) {
        // отделяем сталь от примечаний
        $steel = trim(substr($str, 0, strpos($str, ' ')));
        $comment = trim(substr($str, strpos($str, ' '), strlen($str) - strpos($str, ' ')));
    } else {
        $steel = $str;
    }

    if (strpos($steel, '+N') !== false) { // сталь с обработкой
        $steel = str_replace('+N', '', $steel);
        $processing = 'N';
    } elseif (strpos($steel, '+AR') !== false) { // сталь с обработкой
        $steel = str_replace('+AR', '', $steel);
        $processing = 'AR';
    } else { // без обработки
        $processing = '';
    }

    // записываем данные в объект
    $item->steel = $steel;
    $item->processing = $processing;
    $item->comment = $comment;
}

/**
 * вспомогательная функция для разбора длины
 *
 * принимает на вход строку с содержимым для разбора
 * в строке может быть одно или два числа через дефис
 * функция нужна для сокращения количества кода в функции разбора
 *
 * string    str        строка текста
 *
 * return    string    длина
 */
function parseLength($str) {
    if (strpos($str, '-') !== false) {
        return trim(str_replace(' ', '', $str));
    } else {
        return (int)trim(str_replace(' ', '', $str));
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceSheets($product) {
    global $prices;

    $x = "{$product->thickness}x{$product->width}x{$product->length}x{$product->steel}";
    $y = "{$product->thickness}x*x{$product->steel}";

    if (isset($prices['SHEETS'][$x])) {
        return $prices['SHEETS'][$x];
    } elseif (isset($prices['SHEETS'][$y])) {
        return $prices['SHEETS'][$y];
    } else {
        // попытка найти цену через диапазон размера
        $findPrice = [];
        foreach ($prices['SHEETS'] as $size => $price) {
            $arTmp = explode('x', $size);
            if ($arTmp[0] == $product->thickness && $arTmp[3] == $product->steel
                && $product->width >= $arTmp[1] && $product->length <= $arTmp[2]) {
                $findPrice[$price] = [
                    'thickness' => $arTmp[0],
                    'width'     => $arTmp[1],
                    'length'    => $arTmp[2],
                    'steel'     => $arTmp[3],
                    'price'     => $price
                ];
            }
        }
        if (count($findPrice) > 0) {
            $last = end($findPrice);

            return $last !== false ? $last['price'] : 0;
        }

        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceTeardropPlates($product) {
    global $prices;
    $x = "{$product->thickness}x{$product->width}x{$product->length}x{$product->steel}";
    if (isset($prices['TEARDROP PLATES'][$x])) {
        return $prices['TEARDROP PLATES'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceBoilerPlates($product) {
    global $prices;
    $x = "{$product->thickness}x{$product->width}x{$product->length}x{$product->steel}";
    if (isset($prices['BOILER PLATES'][$x])) {
        return $prices['BOILER PLATES'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceColdRolledPlates($product) {
    global $prices;
    $x = "{$product->thickness}x{$product->width}x{$product->length}x{$product->steel}";
    if (isset($prices['COLD ROLLED PLATES'][$x])) {
        return $prices['COLD ROLLED PLATES'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceLowTBars($product) {
    global $prices;
    $x = "{$product->size}x{$product->steel}";
    if (isset($prices['LOW T-BARS'][$x])) {
        return $prices['LOW T-BARS'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 * @param $product
 * @return int
 */
function getPriceColdFormedChannels($product) {
    global $prices;
    $x = "{$product->size}x{$product->steel}";
    if (isset($prices['COLD FORMED CHANNELS'][$x])) {
        return $prices['COLD FORMED CHANNELS'][$x];
    } else {
        return 0;
    }
}

function getPriceGateProfiles($product) {
    global $prices;
    $x = "{$product->size}x{$product->steel}";
    if (isset($prices['GATE PROFILES'][$x])) {
        return $prices['GATE PROFILES'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceTBars($product) {
    global $prices;
    $x = "{$product->size}x{$product->steel}";
    $y = "{$product->width}x{$product->steel}";
    if (isset($prices['T-BARS'][$x])) {
        return $prices['T-BARS'][$x];
    } elseif (isset($prices['T-BARS'][$y])) {
        return $prices['T-BARS'][$y];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 * @param $product
 * @return int
 */
function getPriceProfil($product) {
    global $prices;
    $product->size = strtolower($product->size); // т.к. массив размеров сделал через "х"
    if ($product->type == 'UPN' && strpos($product->size, 'E') !== false) {
        $type = 'UPE';
        $size = str_replace('E', '', $product->size);
    } else {
        $type = $product->type;
        $size = $product->size;
    }
    $x = "{$type} {$size}x{$product->steel}x{$product->length}";
    $y = "{$type} {$size}x{$product->steel}";
    if (isset($prices['PROFIL'][$x]) && ($prices['PROFIL'][$x])) {
        return $prices['PROFIL'][$x];
    } elseif (isset($prices['PROFIL'][$y])) {
        return $prices['PROFIL'][$y];
    } else {
        $arTmp = explode('x', $size);
        $count = count($arTmp);
        if ($count == 3) {
            $size1 = "{$type} {$arTmp[0]}x{$arTmp[1]}x$product->steel";
            $size2 = "{$type} {$arTmp[0]}x$product->steel";

            return $prices['PROFIL'][$size1] ?? $prices['PROFIL'][$size2] ?? 0;
        } elseif ($count == 2) {
            $size = "{$arTmp[0]}x$product->steel";

            return $prices['PROFIL'][$size] ?? 0;
        }

        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceEqualAngles($product) {
    global $prices;

    $x = "{$product->width}X{$product->thickness}x{$product->steel}";
    $y = "{$product->width}X{$product->height}X{$product->thickness}x{$product->steel}";
    $z = "{$product->width}x{$product->steel}";

    if (isset($prices['EQUAL-LEG ANGLES'][$x])) {
        return $prices['EQUAL-LEG ANGLES'][$x];
    } elseif (isset($prices['EQUAL-LEG ANGLES'][$y])) {
        return $prices['EQUAL-LEG ANGLES'][$y];
    } elseif (isset($prices['EQUAL-LEG ANGLES'][$z])) {
        return $prices['EQUAL-LEG ANGLES'][$z];
    }

    return 0;
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceUnequalAngles($product) {
    global $prices;

    $x = "{$product->width}X{$product->thickness}x{$product->steel}";
    $y = "{$product->width}X{$product->height}X{$product->thickness}X{$product->length}x{$product->steel}";
    $z = "{$product->width}X{$product->height}X{$product->thickness}x{$product->steel}";
    $q = "{$product->width}X{$product->height}x{$product->steel}";
    if (isset($prices['UNEQUAL ANGLES'][$x])) {
        return $prices['UNEQUAL ANGLES'][$x];
    } elseif (isset($prices['UNEQUAL ANGLES'][$y])) {
        return $prices['UNEQUAL ANGLES'][$y];
    } elseif (isset($prices['UNEQUAL ANGLES'][$z])) {
        return $prices['UNEQUAL ANGLES'][$z];
    } elseif (isset($prices['UNEQUAL ANGLES'][$q])) {
        return $prices['UNEQUAL ANGLES'][$q];
    } else {
        return 0;
    }
}

function getPriceOcfpUnequalAngles($product) {
    global $prices;

    $x = "{$product->width}X{$product->height}X{$product->thickness}x{$product->steel}";
    if (isset($prices['OCFP UNEQUAL ANGLES'][$x])) {
        return $prices['OCFP UNEQUAL ANGLES'][$x];
    } else {
        return 0;
    }
}

function getPriceOcfpEqualAngles($product) {
    global $prices;

    $x = "{$product->width}X{$product->height}X{$product->thickness}x{$product->steel}";
    if (isset($prices['OCFP EQUAL ANGLES'][$x])) {
        return $prices['OCFP EQUAL ANGLES'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceHollowSections($product) {
    global $prices;

    if (strpos($product->comment, '10219') !== false) {
        $wg = 'EN10219';
    } elseif (strpos($product->comment, '10210') !== false) {
        $wg = 'EN10210';
    } else {
        $wg = '';
    }
    $x = "{$product->size}x{$wg}x{$product->steel}";
    $y = "{$product->width}x{$product->thickness}x{$wg}x{$product->steel}";
    if (isset($prices['HOLLOW SECTIONS'][$x])) {
        return $prices['HOLLOW SECTIONS'][$x];
    } elseif (isset($prices['HOLLOW SECTIONS'][$y])) {
        return $prices['HOLLOW SECTIONS'][$y];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceFlatBars($product) {
    global $prices;

//    $x = "{$product->width}x{$product->height}x{$product->steel}";
    $x = "{$product->width}x{$product->wall}x{$product->steel}";
    if (isset($prices['FLAT BARS'][$x])) {
        return $prices['FLAT BARS'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceRoundBars($product) {
    global $prices;

    $x = "{$product->diameter}x{$product->steel}";
    if (isset($prices['ROUND BARS'][$x])) {
        return $prices['ROUND BARS'][$x];
    } else {
        return 0;
    }
}

function getPriceDrawnRoundBars($product) {
    global $prices;

    $x = "{$product->diameter}x{$product->steel}";
    if (isset($prices['DRAWN ROUND BARS'][$x])) {
        return $prices['DRAWN ROUND BARS'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceSquareBars($product) {
    global $prices;
    $x = "{$product->width}x{$product->steel}";
    if (isset($prices['SQUARE BARS'][$x])) {
        return $prices['SQUARE BARS'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPriceRibbedBars($product) {
    global $prices;
    $x = "{$product->diameter}";
    if (isset($prices['RIBBED BARS'][$x])) {
        return $prices['RIBBED BARS'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * @param object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 * @return int
 */
function getPriceWireRod($product) {
    global $prices;
    $x = "{$product->diameter}";
    if (isset($prices['WIRE ROD'][$x])) {
        return $prices['WIRE ROD'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход объект товара
 * находит по данным товара нужную цену и отдает ее
 *
 * object    product        объект товара
 *
 * return    float            цена (0, если не найдена)
 */
function getPricePipes($product) {
    global $prices;
    if (strpos($product->comment, '10210') !== false) {
        $wg = 'EN10210';
    } else {
        $wg = 'EN10219';
    }
    $x = "{$product->diameter}x{$wg}x{$product->steel}";
    if (isset($prices['PIPES'][$x])) {
        return $prices['PIPES'][$x];
    } else {
        return 0;
    }
}

/**
 * функция для получения цены товара
 *
 * принимает на вход наименование товара для цен и объект товара
 * передает управление в специализированную функцию для категории товаров
 *
 * @param string    name        наименование товара для цен
 * @param object    product        объект товара
 *
 * return    float            цена (-1 если спецфункция не найдена)
 * @return float|int
 */
function getPrice($name, $product) {
    $name = str_replace([' ', '-'], '', $name);
    $fn = 'getPrice' . $name;
    if (function_exists($fn)) {
        return (float)$fn($product);
    } else {
        return -1;
    }
}

/**
 * функция для разбора конкретной строки файла
 *
 * принимает на вход данные строки и настройки разбора
 * разбирает строку и возвращает детальные данные товара,
 * включая наименования на трех языках и цену
 *
 * array    rowData    данные строки
 * object    cfg            объект с настройками
 *
 * return    object        данные товара
 */
function parseRow($rowData, $cfg) {
    $item = new stdClass();
    $item->orig = $rowData[0];

    // удаляем długość handlowa
    $rowData[0] = str_replace(['długość handlowa'], '0 mm', $rowData[0]);
    // избавляемся от лишних пробелов
    $rowData[0] = trim($rowData[0]);
    $x = preg_replace('!\s+!', ' ', $rowData[0]);

    // разбор удобнее проводить по наименованию товара
    if (strpos($rowData[0], 'BLACHA G/W') !== false) {
        $item->namePL = 'BLACHA G/W';
        $item->nameRU = getName('BLACHA G/W', 'ru', $cfg->namesMap);
        $item->nameEN = getName('BLACHA G/W', 'en', $cfg->namesMap);
        preg_match('/(BLACHA G\/W)(.*?)mm(.*?)mm(.*?)mm(.*)/', $x, $matches);
        $item->thickness = (float)trim(str_replace(',', '.', $matches[2]));
        $item->width = (int)trim(str_replace(' ', '', $matches[3]));
        $item->length = (int)trim(str_replace(' ', '', $matches[4]));
        parseAdditional($matches[5], $item);
        if ($item->steel == 'S235JR' || $item->steel == 'S235JRC'
            || $item->steel == 'S355J2' || $item->steel == 'S355J2C'
            || $item->steel == 'S355MC' || $item->steel == 'S355JR') {
            $item->name1C = str_replace('.', ',', "Лист {$item->thickness}х{$item->width}х{$item->length}мм EN 10025-2");
//            $item->nom = "EN 10025-2: листы стальные горячекатаные из нелегированной конструкционной стали";
            $item->nom = "EN 10025-2: листы г/к из нелегированных конструкционных сталей";
            if ($item->steel == 'S355MC') {
                $item->nom = "EN 10149-2: листы г/к из катаной стали подвергнутой термомеханической обработке";
            }
        }
        $item->price = getPrice('SHEETS', $item);
    } elseif (strpos($rowData[0], 'BLACHA ŁEZKOWA') !== false) {
        $item->namePL = 'BLACHA ŁEZKOWA';
        $item->nameRU = getName('BLACHA ŁEZKOWA', 'ru', $cfg->namesMap);
        $item->nameEN = getName('BLACHA ŁEZKOWA', 'en', $cfg->namesMap);
        preg_match('/(BLACHA ŁEZKOWA)(.*?)mm(.*?)mm(.*?)mm(.*)/', $x, $matches);
        $item->thickness = (float)trim(str_replace(',', '.', $matches[2]));
        $item->width = (int)trim(str_replace(' ', '', $matches[3]));
        $item->length = parseLength($matches[4]);
        parseAdditional($matches[5], $item);
        $item->price = getPrice('TEARDROP PLATES', $item);
    } elseif (strpos($rowData[0], 'BLACHA OCYNKOWANA W KRĘGACH') !== false) {
        $item->namePL = 'BLACHA OCYNKOWANA W KRĘGACH';
        $item->nameRU = getName('BLACHA OCYNKOWANA W KRĘGACH', 'ru', $cfg->namesMap);
        $item->nameEN = getName('BLACHA OCYNKOWANA W KRĘGACH', 'en', $cfg->namesMap);
        preg_match('/(BLACHA OCYNKOWANA W KRĘGACH)(.*?)mm(.*?)mm(.*)/', $x, $matches);
        $item->thickness = (float)trim(str_replace(',', '.', $matches[2]));
        $item->width = (int)trim(str_replace(' ', '', $matches[3]));
        parseAdditional($matches[4], $item);
        $processingPos = strpos($item->steel, '-');
        if ($processingPos !== false) {
            $item->processing = substr($item->steel, $processingPos + 1);
            $item->steel = substr($item->steel, 0, $processingPos);
        }
        $item->price = 0; //getPrice('HDG COIL', $item);
    } elseif (strpos($rowData[0], 'BLACHA OCYNKOWANA') !== false) {
        $item->namePL = 'BLACHA OCYNKOWANA';
        $item->nameRU = getName('BLACHA OCYNKOWANA', 'ru', $cfg->namesMap);
        $item->nameEN = getName('BLACHA OCYNKOWANA', 'en', $cfg->namesMap);
        preg_match('/(BLACHA OCYNKOWANA)(.*?)mm(.*?)mm(.*?)mm(.*)/', $x, $matches);
        $item->thickness = (float)trim(str_replace(',', '.', $matches[2]));
        $item->width = (int)trim(str_replace(' ', '', $matches[3]));
        $item->length = parseLength($matches[4]);
        parseAdditional($matches[5], $item);
        $processingPos = strpos($item->steel, '-');
        if ($processingPos !== false) {
            $item->processing = substr($item->steel, $processingPos + 1);
            $item->steel = substr($item->steel, 0, $processingPos);
        }
        $item->price = 0; //getPrice('HDG SHEET', $item);
    } elseif (strpos($rowData[0], 'BLACHA KOTŁOWA') !== false) {
        $item->namePL = 'BLACHA KOTŁOWA';
        $item->nameRU = getName('BLACHA KOTŁOWA', 'ru', $cfg->namesMap);
        $item->nameEN = getName('BLACHA KOTŁOWA', 'en', $cfg->namesMap);
        preg_match('/(BLACHA KOTŁOWA)(.*?)mm(.*?)mm(.*?)mm(.*)/', $x, $matches);
        $item->thickness = (float)trim(str_replace(',', '.', $matches[2]));
        $item->width = (int)trim(str_replace(' ', '', $matches[3]));
        $item->length = parseLength($matches[4]);
        parseAdditional($matches[5], $item);
        $item->price = getPrice('BOILER PLATES', $item);
    } elseif (strpos($rowData[0], 'BLACHA Z/W') !== false) {
        $item->namePL = 'BLACHA Z/W';
        $item->nameRU = getName('BLACHA Z/W', 'ru', $cfg->namesMap);
        $item->nameEN = getName('BLACHA Z/W', 'en', $cfg->namesMap);
        preg_match('/(BLACHA Z\/W)(.*?)mm(.*?)mm(.*?)mm(.*)/', $x, $matches);
        $item->thickness = (float)trim(str_replace(',', '.', $matches[2]));
        $item->width = (int)trim(str_replace(' ', '', $matches[3]));
        $item->length = parseLength($matches[4]);
        parseAdditional($matches[5], $item);

        if ($item->steel == 'DC01') {
            $item->name1C = str_replace('.', ',', "Лист {$item->thickness}х{$item->width}х{$item->length}мм EN 10130");
            $item->nom = "EN 10130: листы х/к из низкоуглеродистой стали для холодной штамповки";
        }

        $item->price = getPrice('COLD ROLLED PLATES', $item);
    } elseif (strpos($rowData[0], 'CEOWNIK EKON. O PÓŁKACH RÓWNOLEG.') !== false) {
        $item->namePL = 'CEOWNIK EKON. O PÓŁKACH RÓWNOLEG.';
        $item->nameRU = getName('CEOWNIK EKON. O PÓŁKACH RÓWNOLEG.', 'ru', $cfg->namesMap);
        $item->nameEN = getName('CEOWNIK EKON. O PÓŁKACH RÓWNOLEG.', 'en', $cfg->namesMap);
        preg_match('/(CEOWNIK EKON. O PÓŁKACH RÓWNOLEG.) (UPE) (\S+) (.*?)mm(.*)/', $x, $matches);
        $item->type = trim($matches[2]);
        $item->size = (int)trim($matches[3]);
        $item->length = parseLength($matches[4]);
        parseAdditional($matches[5], $item);
        $item->name1C = str_replace('.', ',', "Швеллер {$item->type} {$item->size}");
        $item->nom = "Швеллер {$item->type}";
        $item->price = getPrice('PROFIL', $item);
    } elseif (strpos($rowData[0], 'CEOWNIK PÓŁZAMKNIĘTY') !== false) {
        $item->namePL = 'CEOWNIK PÓŁZAMKNIĘTY';
        $item->nameRU = getName('CEOWNIK PÓŁZAMKNIĘTY', 'ru', $cfg->namesMap);
        $item->nameEN = getName('CEOWNIK PÓŁZAMKNIĘTY', 'en', $cfg->namesMap);
        preg_match('/(CEOWNIK PÓŁZAMKNIĘTY) (\S+) (.*?)mm(.*)/', $x, $matches);
        $item->size = trim($matches[2]);
        $item->length = parseLength($matches[3]);
        parseAdditional($matches[4], $item);
        $item->price = getPrice('GATE PROFILES', $item);
    } elseif (strpos($rowData[0], 'CEOWNIK Z/G') !== false) {
        $item->namePL = 'CEOWNIK Z/G';
        $item->nameRU = getName('CEOWNIK Z/G', 'ru', $cfg->namesMap);
        $item->nameEN = getName('CEOWNIK Z/G', 'en', $cfg->namesMap);
        preg_match('/(CEOWNIK) (UPN|Z\/G) (\S+) (.*?)mm(.*)/', $x, $matches);
        $item->type = trim($matches[2]);
        $item->size = trim($matches[3]);
        $item->length = parseLength($matches[4]);
        if ($item->type == 'Z/G') {
            list($item->width, $item->height, $item->thickness) = explode('x', strtolower($item->size));
            $item->thickness = (float)str_replace(',', '.', $item->thickness);
        }
        parseAdditional($matches[5], $item);
        $item->price = ($item->type == 'Z/G') ? getPrice('COLD FORMED CHANNELS', $item) : getPrice('PROFIL', $item);
    } elseif (strpos($rowData[0], 'CEOWNIK') !== false) {
        $item->namePL = 'CEOWNIK';
        $item->nameRU = getName('CEOWNIK', 'ru', $cfg->namesMap);
        $item->nameEN = getName('CEOWNIK', 'en', $cfg->namesMap);
        preg_match('/(CEOWNIK) (UPN|Z\/G) (\S+) (.*?)mm(.*)/', $x, $matches);
        $item->type = trim($matches[2]);
        $item->size = trim($matches[3]);
        $item->length = parseLength($matches[4]);
        if ($item->type == 'Z/G') {
            list($item->width, $item->height, $item->thickness) = explode('x', strtolower($item->size));
            $item->thickness = (float)str_replace(',', '.', $item->thickness);
        }
        parseAdditional($matches[5], $item);
        $tmpSize = str_replace('X', 'х', $item->size);
        $item->name1C = str_replace('.', ',', "Швеллер {$item->type} {$tmpSize}");
        if ($item->type == 'UPN') {
            if (strpos($item->size, 'X') !== false) {
                $item->name1C .= ' (U)';
            } else {
                $item->name1C .= ' (UNP)';
            }
        }
        $item->nom = "Швеллер {$item->type}";

        if ($item->type == 'UPN' && in_array(strtolower($item->size), ['100e', '120e', '140e', '160e', '180e', '80e'])) {
            $item->name1C = '';
            $item->nom = '';
        }

        $item->price = ($item->type == 'Z/G') ? 0 : getPrice('PROFIL', $item);
    } elseif (strpos($rowData[0], 'DWUTEOWNIK') !== false) {
        $item->namePL = 'DWUTEOWNIK';
        $item->nameRU = getName('DWUTEOWNIK', 'ru', $cfg->namesMap);
        $item->nameEN = getName('DWUTEOWNIK', 'en', $cfg->namesMap);
        preg_match('/(DWUTEOWNIK) (HEA|HEB|HEM|IPE|IPN) (\S+) (.*?)mm(.*)/', $x, $matches);
        $item->type = trim($matches[2]);
        $item->size = trim($matches[3]);
        $item->length = parseLength($matches[4]);
        parseAdditional($matches[5], $item);
        if ($item->type == 'HEA') {
            $item->name1C = str_replace('.', ',', "Балка HEA {$item->size} (IPB l)");
            $item->nom = "Балка HEA (IPB l)";
        } elseif ($item->type == 'HEB') {
            $item->name1C = str_replace('.', ',', "Балка HEB {$item->size} (IPB)");
            $item->nom = "Балка HEB (IPB)";
        } elseif ($item->type == 'HEM') {
            $item->name1C = str_replace('.', ',', "Балка HEM {$item->size} (IPB v)");
            $item->nom = "Балка HEM (IPB v)";
        } elseif ($item->type == 'IPE') {
            $item->name1C = str_replace('.', ',', "Балка IPE {$item->size}");
            $item->nom = "Балка IPE";
        } elseif ($item->type == 'IPN') {
            $item->name1C = str_replace('.', ',', "Балка IPN {$item->size} (INP)");
            $item->nom = "Балка IPN (INP)";
        }
        $item->price = getPrice('PROFIL', $item);
    } elseif (strpos($rowData[0], 'KĄTOWNIK Z/G') !== false) {
        $item->namePL = 'KĄTOWNIK Z/G';
        $item->nameRU = getName('KĄTOWNIK Z/G', 'ru', $cfg->namesMap);
        $item->nameEN = getName('KĄTOWNIK Z/G', 'en', $cfg->namesMap);
        preg_match('/(KĄTOWNIK) (Z\/G )?(\S+) (.*?)mm(.*)/', $x, $matches);
        $item->type = trim($matches[2]);
        $item->size = trim($matches[3]);
        $item->length = parseLength($matches[4]);
        list($item->width, $item->height, $item->thickness) = explode('X', $item->size);
        $item->thickness = (float)str_replace(',', '.', $item->thickness);
        parseAdditional($matches[5], $item);
        $item->price = ($item->type == 'Z/G')
            ? getPrice($item->width == $item->height ? 'OCFP EQUAL ANGLES' : 'OCFP UNEQUAL ANGLES', $item)
            : getPrice($item->width == $item->height ? 'EQUAL ANGLES' : 'UNEQUAL ANGLES', $item);
    } elseif (strpos($rowData[0], 'KĄTOWNIK') !== false) {
        $item->namePL = 'KĄTOWNIK';
        $item->nameRU = getName('KĄTOWNIK', 'ru', $cfg->namesMap);
        $item->nameEN = getName('KĄTOWNIK', 'en', $cfg->namesMap);
        preg_match('/(KĄTOWNIK) (Z\/G )?(\S+) (.*?)mm(.*)/', $x, $matches);
        $item->type = trim($matches[2]);
        $item->size = trim($matches[3]);
        $item->length = parseLength($matches[4]);
        list($item->width, $item->height, $item->thickness) = explode('X', $item->size);
        $item->thickness = (float)str_replace(',', '.', $item->thickness);
        parseAdditional($matches[5], $item);
        if ($item->width == $item->height) {
            $item->name1C = str_replace('.', ',', "Уголок равнополочный {$item->width}х{$item->height}х{$item->thickness} EN 10056");
            $item->nom = "EN 10056: уголки равнополочные (равнобедренные)";
        } else {
            $item->name1C = str_replace('.', ',', "Уголок неравнополочный {$item->width}х{$item->height}х{$item->thickness} EN 10056");
            $item->nom = "EN 10056: уголки неравнополочные (неравнобедренные)";
        }
        $item->price = ($item->type == 'Z/G')
            ? getPrice($item->width == $item->height ? 'OCFP EQUAL ANGLES' : 'OCFP UNEQUAL ANGLES', $item)
            : getPrice($item->width == $item->height ? 'EQUAL ANGLES' : 'UNEQUAL ANGLES', $item);
    } elseif (strpos($rowData[0], 'KSZTAŁTOWNIK KWADRATOWY') !== false) {
        $item->namePL = 'KSZTAŁTOWNIK KWADRATOWY';
        $item->nameRU = getName('KSZTAŁTOWNIK KWADRATOWY', 'ru', $cfg->namesMap);
        $item->nameEN = getName('KSZTAŁTOWNIK KWADRATOWY', 'en', $cfg->namesMap);
        preg_match('/(KSZTAŁTOWNIK KWADRATOWY) (G\/W )?(\S+) (.*?)mm(.*)/', $x, $matches);
        $item->type = trim($matches[2]);
        $item->size = trim($matches[3]);
        $item->length = parseLength($matches[4]);
        list($item->width, $item->height, $item->thickness) = explode('X', $item->size);
        $item->thickness = (float)str_replace(',', '.', $item->thickness);
        parseAdditional($matches[5], $item);
        if ($item->steel == 'E235') {
            $item->name1C = str_replace('.', ',', "Труба квадратная сварная {$item->width}х{$item->height}х{$item->thickness} EN 10305-5");
            $item->nom = "EN 10305-5: трубы квадратные прецизионные сварные холоднокалиброванные";
        } elseif ($item->steel == 'S355J2H' || $item->steel == 'S355J2H/S420MH'
            || $item->steel == 'S235JRH' || $item->steel == 'S235JR' || $item->steel == 'S275J0H') {
            if ($item->type) {
                $item->name1C = str_replace('.', ',', "Труба квадратная бесшовная {$item->width}х{$item->height}х{$item->thickness} EN 10210");
                $item->nom = "EN 10210: трубы квадратные бесшовные горячекатаные";
            } else {
                $item->name1C = str_replace('.', ',', "Труба квадратная {$item->width}х{$item->height}х{$item->thickness} EN 10219");
                $item->nom = "EN 10219: трубы квадратные сварные холоднокатаные";
            }
        }
        $item->price = getPrice('HOLLOW SECTIONS', $item);
    } elseif (strpos($rowData[0], 'KSZTAŁTOWNIK PROSTOKĄTNY') !== false) {
        $item->namePL = 'KSZTAŁTOWNIK PROSTOKĄTNY';
        $item->nameRU = getName('KSZTAŁTOWNIK PROSTOKĄTNY', 'ru', $cfg->namesMap);
        $item->nameEN = getName('KSZTAŁTOWNIK PROSTOKĄTNY', 'en', $cfg->namesMap);
        preg_match('/(KSZTAŁTOWNIK PROSTOKĄTNY) (G\/W )?(\S+) (.*?)mm(.*)/', $x, $matches);
        $item->type = trim($matches[2]);
        $item->size = trim($matches[3]);
        $item->length = parseLength($matches[4]);
        list($item->width, $item->height, $item->thickness) = explode('X', $item->size);
        $item->thickness = (float)str_replace(',', '.', $item->thickness);
        parseAdditional($matches[5], $item);
        if ($item->steel == 'E235') {
            $item->name1C = str_replace('.', ',', "Труба прямоугольная сварная {$item->width}х{$item->height}х{$item->thickness} EN 10305-5");
            $item->nom = "EN 10305-5: трубы прямоугольные прецизионные сварные холоднокалиброванные";
        } elseif ($item->steel == 'S355J2H'
            || $item->steel == 'S235JRH' || $item->steel == 'S355J2H/S420MH'
            || $item->steel == 'S235JR' || $item->steel == 'S275J2H') {
            if ($item->type) {
                $item->name1C = str_replace('.', ',', "Труба прямоугольная бесшовная {$item->width}х{$item->height}х{$item->thickness} EN 10210");
                $item->nom = "EN 10210: трубы прямоугольные бесшовные горячекатаные";
            } else {
                $item->name1C = str_replace('.', ',', "Труба прямоугольная {$item->width}х{$item->height}х{$item->thickness} EN 10219");
                $item->nom = "EN 10219: трубы прямоугольные сварные холоднокатаные";
            }
        }
        $item->price = getPrice('HOLLOW SECTIONS', $item);
    } elseif (strpos($rowData[0], 'PŁASKOWNIK') !== false) {
        $item->namePL = 'PŁASKOWNIK';
        $item->nameRU = getName('PŁASKOWNIK', 'ru', $cfg->namesMap);
        $item->nameEN = getName('PŁASKOWNIK', 'en', $cfg->namesMap);
        preg_match('/(PŁASKOWNIK) (\S+) (.*?)mm(.*)/', $x, $matches);
        $item->size = trim($matches[2]);
        $item->length = parseLength($matches[3]);
//        list($item->width, $item->height) = explode('X', $item->size);
        list($item->width, $item->wall) = explode('X', $item->size);
        $item->thickness = $item->wall;
        parseAdditional($matches[4], $item);
        if ($item->steel != 'ST52-3') {
//            $item->name1C = str_replace('.', ',', "Полоса {$item->height}х{$item->width}мм EN 10025-2");
//            $item->name1C = str_replace('.', ',', "Полоса {$item->wall}х{$item->width}мм EN 10025-2");
            $item->name1C = str_replace('.', ',', "Полоса {$item->thickness}х{$item->width}мм EN 10025-2");
            $item->nom = "EN 10025-2: полосы стальные горячекатаные из нелегированной конструкционной стали";
        }
        $item->price = $item->steel != 'ST52-3' ? getPrice('FLAT BARS', $item) : 0;
        unset($item->wall);
    } elseif (strpos($rowData[0], 'PRĘT GŁADKI CIĄGNIONY') !== false) {
        $item->namePL = 'PRĘT GŁADKI CIĄGNIONY';
        $item->nameRU = getName('PRĘT GŁADKI CIĄGNIONY', 'ru', $cfg->namesMap);
        $item->nameEN = getName('PRĘT GŁADKI CIĄGNIONY', 'en', $cfg->namesMap);
        preg_match('/(PRĘT GŁADKI CIĄGNIONY) (fi )?(\S+) (.*?)mm(.*)/', $x, $matches);
        $item->diameter = trim($matches[3]);
        $item->length = trim(str_replace(' ', '', $matches[4]));
        parseAdditional($matches[5], $item);
        if ($item->steel == '11SMn30') {
            $item->name1C = str_replace('.', ',', "Круг (пруток) {$item->diameter}мм EN 10277-3");
            $item->nom = "EN 10277-3: круги стальные автоматные, сталь повышенной отделки поверхности";
        } elseif ($item->steel == 'C45') {
            $item->name1C = str_replace('.', ',', "Круг (пруток) {$item->diameter}мм EN 10277-2");
            $item->nom = "EN 10277-2: круги стальные для общего машиностроения, сталь повышенной отделки поверхности";
        } elseif ($item->steel == 'S235JR' || $item->steel == 'S355J2') {
            $item->name1C = str_replace('.', ',', "Круг (пруток) {$item->diameter}мм EN 10025-2");
            $item->nom = "EN 10025-2: круги стальные горячекатаные из нелегированной конструкционной стали";
        }
        $item->price = getPrice('DRAWN ROUND BARS', $item);
    } elseif (strpos($rowData[0], 'PRĘT GŁADKI') !== false) {
        $item->namePL = 'PRĘT GŁADKI';
        $item->nameRU = getName('PRĘT GŁADKI', 'ru', $cfg->namesMap);
        $item->nameEN = getName('PRĘT GŁADKI', 'en', $cfg->namesMap);
        preg_match('/(PRĘT GŁADKI) (fi )?(\S+) (.*?)mm(.*)/', $x, $matches);
        $item->diameter = (float)str_replace(',', '.', $matches[3]);
        $item->length = parseLength($matches[4]);
        parseAdditional($matches[5], $item);
        if ($item->steel == '11SMn30') {
            $item->name1C = str_replace('.', ',', "Круг (пруток) {$item->diameter}мм EN 10277-3");
            $item->nom = "EN 10277-3: круги стальные автоматные, сталь повышенной отделки поверхности";
        } elseif ($item->steel == 'C45' || $item->steel == 'C45R') {
            $item->name1C = str_replace('.', ',', "Круг (пруток) {$item->diameter}мм EN 10277-2");
            $item->nom = "EN 10277-2: круги стальные для общего машиностроения, сталь повышенной отделки поверхности";
        } elseif ($item->steel == 'C45E') {
            $item->name1C = str_replace('.', ',', "Круг (пруток) {$item->diameter}мм EN 10277-5");
            $item->nom = "EN 10277-5: круги стальные для закалки и отпуска, сталь повышенной отделки поверхности";
        } elseif ($item->steel == '41Cr4') {
            $item->name1C = str_replace('.', ',', "Круг (пруток) {$item->diameter}мм EN 10083-3");
            $item->nom = "EN 10083-3: круги стальные из легированной стали для закаливания и отпуска";
        } elseif ($item->steel == 'S235JR' || $item->steel == 'S355J2' || $item->steel == 'S235JR/S275JR') {
            $item->name1C = str_replace('.', ',', "Круг (пруток) {$item->diameter}мм EN 10025-2");
            $item->nom = "EN 10025-2: круги стальные горячекатаные из нелегированной конструкционной стали";
        }
        $item->price = getPrice('ROUND BARS', $item);
    } elseif (strpos($rowData[0], 'PRĘT KWADRATOWY') !== false) {
        $item->namePL = 'PRĘT KWADRATOWY';
        $item->nameRU = getName('PRĘT KWADRATOWY', 'ru', $cfg->namesMap);
        $item->nameEN = getName('PRĘT KWADRATOWY', 'en', $cfg->namesMap);
        preg_match('/(PRĘT KWADRATOWY) (\S+) (.*?)mm(.*)/', $x, $matches);
        $item->size = trim($matches[2]);
        list($item->width, $item->height) = explode('X', $item->size);
        $item->length = parseLength($matches[3]);
        parseAdditional($matches[4], $item);
        if ($item->steel == 'C45') {
            $item->name1C = str_replace('.', ',', "Квадрат стальной {$item->height}мм EN 10277-2");
            $item->nom = "EN 10277-2: квадраты стальные для общего машиностроения, сталь повышенной отделки поверхности";
        } elseif (in_array($item->steel, ['S355J2', 'S275JR', 'S235JR', 'S235JR/S275JR'])) {
            $item->name1C = str_replace('.', ',', "Квадрат стальной {$item->height}мм EN 10025-2");
            $item->nom = "EN 10025-2: квадраты стальные горячекатаные из нелегированной конструкционной стали";
        }
        $item->price = getPrice('SQUARE BARS', $item);
    } elseif (strpos($rowData[0], 'PRĘT ŻEBROWANY') !== false) {
        $item->namePL = 'PRĘT ŻEBROWANY';
        $item->nameRU = getName('PRĘT ŻEBROWANY', 'ru', $cfg->namesMap);
        $item->nameEN = getName('PRĘT ŻEBROWANY', 'en', $cfg->namesMap);
        preg_match('/(PRĘT ŻEBROWANY) (fi )?(\S+) (.*?)mm(.*)/', $x, $matches);
        $item->diameter = (float)str_replace(',', '.', $matches[3]);
        $item->length = parseLength($matches[4]);
        parseAdditional($matches[5], $item);

        if ($item->steel == 'B500B') {
            $item->name1C = str_replace('.', ',', "Квадрат стальной {$item->diameter}мм EN 10225-2");
            $item->nom = "EN 10025-2: квадраты стальные горячекатаные из нелегированной конструкционной стали";
        }

        $item->price = getPrice('RIBBED BARS', $item);
    } elseif (strpos($rowData[0], 'RURA') !== false) {
        $item->namePL = 'RURA';
        $item->nameRU = getName('RURA', 'ru', $cfg->namesMap);
        $item->nameEN = getName('RURA', 'en', $cfg->namesMap);

        if (strpos($x, '"') == false) {
            preg_match('/(RURA) (Z\/SZW )?(\S+) (.*?)mm(.*)/', $x, $matches);

            $item->type = trim($matches[2]);
            $item->size = trim($matches[3]);
            $item->length = parseLength($matches[4]);

            list($item->diameter, $item->wall) = explode('X', $item->size);
            $item->diameter = (float)str_replace(',', '.', $item->diameter);
            $item->wall = (float)str_replace(',', '.', $item->wall);

            parseAdditional($matches[5], $item);
        } else {
            preg_match('/(RURA) (Z\/SZW )?(\S+) (\((.*)\)) (.*?)mm(.*)/', $x, $matches);

            $item->type = trim($matches[2]);
            $item->size = trim($matches[3]);
            $item->inches = trim($matches[5]);
            $item->length = parseLength($matches[6]);

            list($item->diameter, $item->wall) = explode('X', $item->size);
            $item->diameter = (float)str_replace(',', '.', $item->diameter);
            $item->wall = (float)str_replace(',', '.', $item->wall);

            parseAdditional($matches[7], $item);
        }

        if (in_array($item->steel, ['S235JRH', 'S235JR'])) {
            $item->name1C = str_replace('.', ',', "Труба круглая сварная {$item->diameter}х{$item->wall} EN 10219");
            $item->nom = "EN 10219: трубы круглые сварные холоднокатаные";
        }
        $item->price = getPrice('PIPES', $item);
    } elseif (strpos($rowData[0], 'TEOWNIK') !== false) {
        $item->namePL = 'TEOWNIK';
        $item->nameRU = getName('TEOWNIK', 'ru', $cfg->namesMap);
        $item->nameEN = getName('TEOWNIK', 'en', $cfg->namesMap);
        preg_match('/(TEOWNIK) (NISKI )?(\S+) (.*?)mm(.*)/', $x, $matches);
        $item->size = trim($matches[3]);
        list($item->width, $item->height, $item->thickness) = explode('X', $item->size);
        $item->thickness = (float)str_replace(',', '.', $item->thickness);
        $item->length = parseLength($matches[4]);
        parseAdditional($matches[5], $item);
        $item->name1C = str_replace('.', ',', "Балка тавровая Т{$item->height}");
        $item->nom = "Балка Т";
        $item->price = getPrice($matches[2] ? 'LOW T-BARS' : 'T-BARS', $item);
    } elseif (strpos($rowData[0], 'WALCÓWKA') !== false) {
        $item->namePL = 'WALCÓWKA';
        $item->nameRU = getName('WALCÓWKA', 'ru', $cfg->namesMap);
        $item->nameEN = getName('WALCÓWKA', 'en', $cfg->namesMap);
//        preg_match('/(WALCÓWKA) (fi )?(\S+) (.*)/', $x, $matches);
        preg_match('/(WALCÓWKA) (fi )?(\S+)\s+([^\s]+)(.*)/', $x, $matches);
        $item->diameter = (float)str_replace(',', '.', $matches[3]);
        parseAdditional($matches[4], $item);
//        $item->comment = $matches[5] ?? '';1
        $item->price = getPrice('WIRE ROD', $item);
        $item->steel .= ! empty($matches[5]) ? (' ' . $matches[5]) : '';
    } else {
        // неопределенный товар, не подошел ни один шаблон разбора
        $item->namePL = '***UNKNOWN***';
    }

    /*$item->quantity = (float)$rowData[1];
    $item->unit = $rowData[2];*/

    $item->quantity = (float)$rowData[$cfg->titles['quantity']];
    $item->unit = $rowData[$cfg->titles['unit']];

    return $item;
}

/**
 * функция для разбора файла цен
 *
 * принимает на вход путь к файлу цен и настройки разбора
 * разбирает файл и отдает в удобном виде для последующих манипуляций
 *
 * @param string    priceFileName    путь к файлу цен
 * @param object    cfg                    объект с настройками
 *
 * return    array    разобранный массив цен
 * @return array
 */
function parsePrices($priceFileName, $cfg) {
    $data = []; // здесь будут храниться разобранные данные цен
    // подключение утилиты для разбора XLS
    //require_once(dirname(__FILE__) . '/vendor/PHPExcel/IOFactory.php');

    // флаг для определения, нужно ли продолжать разбор
    $processing = true;

    try { // попытка открыть файл на чтение
        $inputFileType = PHPExcel_IOFactory::identify($priceFileName);
        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objPHPExcel = $objReader->load($priceFileName);
    } catch (Exception $e) { // что-то пошло не так (напирмер, битый файл)
        echo "error loading file - " . $e->getMessage();
        $processing = false; // указываем, что продолжать не нужно
    }

    if ($processing) { // файл открылся на чтение, можно разбирать
        // ********************************************************************
        // start SHEETS & PLATES
        // ********************************************************************
        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('SHEETS & PLATES');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        // ***************
        // start SHEETS
        // ***************
        $data['SHEETS'] = \App\PriceParser::getSheets($objPHPExcel);
        // ***************
        // end SHEETS
        // ***************

        // ***************
        // start TEARDROP PLATES
        // ***************

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[0])) == 'TEARDROP PLATES') {
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
            // убираем лишние пробелы
            $sizes = trim(preg_replace('!\s+!', ' ', $rowData[1]));
            // режем строку в массив
            $sizes = explode(' ', $sizes);

            // цена за сталь S235JR
            $steel = trim($rowData[2]);

            // создаем специальный хеш для хранения цен
            foreach ($thicknesses as $t) {
                foreach ($sizes as $s) {
                    $x["{$t}x{$s}xS235JR"] = $steel;
                }
            }
        }
        $data['TEARDROP PLATES'] = $x;

        // ***************
        // end TEARDROP PLATES
        // ***************

        // ***************
        // start BOILER PLATES
        // ***************

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[0])) == 'BOILER PLATES') {
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
            // убираем лишние пробелы
            $sizes = trim(preg_replace('!\s+!', ' ', $rowData[1]));
            // режем строку в массив
            $sizes = explode(' ', $sizes);

            // цена за сталь P235GH
            $steel = trim($rowData[2]);

            // создаем специальный хеш для хранения цен
            foreach ($thicknesses as $t) {
                foreach ($sizes as $s) {
                    $x["{$t}x{$s}xP235GH"] = $steel;
                }
            }
        }
        $data['BOILER PLATES'] = $x;

        // ***************
        // end BOILER PLATES
        // ***************

        // ***************
        // start COLD ROLLED PLATES
        // ***************
        $data['COLD ROLLED PLATES'] = \App\PriceParser::getColdRolledPlates($objPHPExcel);
        // ***************
        // end COLD ROLED PLATES
        // ***************

        // ********************************************************************
        // end SHEETS & PLATES
        // ********************************************************************

        // ********************************************************************
        // start MERCHANT BARS
        // ********************************************************************

        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('Merchant Bars');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();


        $data['EQUAL-LEG ANGLES'] = \App\PriceParser::getEqualLegAngles($objPHPExcel);
        $data['UNEQUAL ANGLES'] = \App\PriceParser::getUnequalAngles($objPHPExcel);

        // ***************
        // start T-BARS
        // ***************
        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[1])) == 'T-BARS') {
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

            // цена за сталь S235JR
            $steel1 = trim($rowData[2]);
            // цена за сталь S355J2
            $steel2 = trim($rowData[3]);

            // создаем специальный хеш для хранения цен
            foreach ($sizes as $s) {
                $x["{$s}xS235JR"] = $steel1;
                $x["{$s}xS355J2"] = $steel2;
            }
        }
        $data['T-BARS'] = $x;

        // ***************
        // end T-BARS
        // ***************

        // ***************
        // start LOW T-BARS
        // ***************

        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[1])) == 'LOW T-BARS') {
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
            // избавляемся от пробелов, приводим к верхнему регистру
            $sizes = strtoupper(str_replace(' ', '', $rowData[1]));
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

            // цена за сталь S235JR
            $steel1 = trim($rowData[2]);
            // цена за сталь S355J2
            $steel2 = trim($rowData[3]);

            // создаем специальный хеш для хранения цен
            foreach ($sizes as $s) {
                $x["{$s}xS235JR"] = $steel1;
                $x["{$s}xS355J2"] = $steel2;
            }
        }
        $data['LOW T-BARS'] = $x;

        // ***************
        // end LOW T-BARS
        // ***************

        // ********************************************************************
        // end MERCHANT BARS
        // ********************************************************************

        // ********************************************************************
        // start OCFP
        // ********************************************************************

        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('OCFP');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        // ***************
        // start COLD FORMED CHANELS
        // ***************
        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[1])) == 'COLD FORMED CHANELS') {
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
            // заменяем запятые на точку с запятой
            // избавляемся от пробелов, приводим к верхнему регистру
            $sizes = strtoupper(str_replace([' ', ','], ['', ';'], $rowData[1]));
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

            // цена за сталь S235JR
            $steel1 = trim($rowData[2]);

            // создаем специальный хеш для хранения цен
            foreach ($sizes as $s) {
                $x["{$s}xS235JR"] = $steel1;
            }
        }
        $data['COLD FORMED CHANNELS'] = $x;

        // ***************
        // end COLD FORMED CHANELS
        // ***************

        // ***************
        // start GATE PROFILES
        // ***************
        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[1])) == 'GATE PROFILES') {
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
            // заменяем запятые на точку с запятой
            // избавляемся от пробелов, приводим к верхнему регистру
            $sizes = strtoupper(str_replace([' ', ','], ['', '.'], $rowData[1]));

            $parts = explode('X', $sizes);
            array_walk($parts, function (&$item) {
                $item = 1 * $item;
            });
            $sizes = implode('X', $parts);
            $sizes = strtoupper(str_replace('.', ',', $sizes));

            // цена за сталь S235JR
            $steel1 = trim($rowData[2]);

            // создаем специальный хеш для хранения цен
            $x["{$sizes}xS235JR"] = $steel1;
        }
        $data['GATE PROFILES'] = $x;

        // ***************
        // end GATE PROFILES
        // ***************

        // ***************
        // start OCFP UNEQUAL ANGLES
        // ***************
        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[1])) == 'UNEQUAL COLD FORMED ANGLES') {
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
            // избавляемся от пробелов, приводим к верхнему регистру
            $sizes = [];
            $str = strtoupper(str_replace([' ', 'z/g'], ['', ''], $rowData[1]));
            preg_match('/(\S+)TO(\S+)X((\d+)-(\d+))/', $str, $matches);
            if (! empty($matches)) {
                $sizeFrom = explode('X', $matches[1]);
                $sizeTo = explode('X', $matches[2]);
                $thicknessFrom = (int)$matches[4];
                $thicknessTo = (int)$matches[5];

                $widths = range($sizeFrom[0], $sizeTo[0]);
                $heights = range($sizeFrom[1], $sizeTo[1]);
                $thicknesses = range($thicknessFrom, $thicknessTo);
                foreach ($thicknesses as $thickness) {
                    foreach ($widths as $width) {
                        foreach ($heights as $height) {
                            $sizes[] = $width . 'X' . $height . 'X' . $thickness;
                        }
                    }
                }
            } else {
                $sizes[] = $str;
            }
            // цена за сталь S235JR
            $steel1 = trim($rowData[2]);

            // создаем специальный хеш для хранения цен
            foreach ($sizes as $s) {
                $x["{$s}xS235JR"] = $steel1;
            }
        }
        $data['OCFP UNEQUAL ANGLES'] = $x;

        // ***************
        // end OCFP UNEQUAL ANGLES
        // ***************

        // ***************
        // start OCFP EQUAL ANGLES
        // ***************
        $x = [];
        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (strtoupper(trim($rowData[1])) == 'COLD FORMED ANGLES') {
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
            // избавляемся от пробелов, приводим к верхнему регистру
            $sizes = [];
            $str = strtoupper(str_replace([' ', 'z/g'], ['', ''], $rowData[1]));
            preg_match('/(\S+)TO(\S+)X((\d+)-(\d+))/', $str, $matches);
            if (! empty($matches)) {
                $sizeFrom = explode('X', $matches[1]);
                $sizeTo = explode('X', $matches[2]);
                $thicknessFrom = (int)$matches[4];
                $thicknessTo = (int)$matches[5];

                $widths = range($sizeFrom[0], $sizeTo[0]);
                $heights = range($sizeFrom[1], $sizeTo[1]);
                $thicknesses = range($thicknessFrom, $thicknessTo);
                foreach ($thicknesses as $thickness) {
                    foreach ($widths as $width) {
                        foreach ($heights as $height) {
                            $sizes[] = $width . 'X' . $height . 'X' . $thickness;
                        }
                    }
                }
            } else {
                $sizes[] = $str;
            }
            // цена за сталь S235JR
            $steel1 = trim($rowData[2]);

            // создаем специальный хеш для хранения цен
            foreach ($sizes as $s) {
                $x["{$s}xS235JR"] = $steel1;
            }
        }
        $data['OCFP EQUAL ANGLES'] = $x;

        // ***************
        // end OCFP UNEQUAL ANGLES
        // ***************

        // ********************************************************************
        // end OCFP
        // ********************************************************************

        // ********************************************************************
        // start PRICES
        // ********************************************************************
//        $data['PROFIL'] = parsePriceSheet($objPHPExcel, $cfg);
        $data['PROFIL'] = \App\PriceParser::getPriceSheet($objPHPExcel, $cfg);
        // ********************************************************************
        // end PRICES
        // ********************************************************************

        // ********************************************************************
        // start HOLLOW SECTIONS
        // ********************************************************************

        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('Hollow Sections ');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        // ********************************************************************
        // start PIPES
        // ********************************************************************
        $x = [];

        // флаг, который определяет начало данных
        $startData = false;
        // цикл разбора, от стартовой строки до нижней
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
            if (trim($rowData[1]) == 'construction pipes acc  EN 10219') {
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

            // убираем лишние пробелы и фразы, заменяем запятую точкой
            $s = str_replace(['mm', ' ', ','], ['', '', '.'], trim($rowData[1]));
            // получаем диапазон диаметров
            list($a, $b) = explode('to', $s);
            $diameters = range($a, $b + 0.1, 0.1);

            // цена за сталь S235JR(H), E235
            $steel = trim($rowData[2]);

            // создаем специальный хеш для хранения цен
            foreach ($diameters as $d) {
                $x["{$d}xEN10219xS235JR"] = $steel;
                $x["{$d}xEN10219xS235JRH"] = $steel;
                $x["{$d}xEN10219xE235"] = $steel;
            }
        }

        $data['PIPES'] = $x;

        // ********************************************************************
        // end PIPES
        // ********************************************************************

        $x = \App\PriceParser::getColdFormedHollowSections($objPHPExcel);
        /*$x = [];
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
        }*/

        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('Hollow Sections EN10210');

        if ($sheet) {
            // определяем самую нижнюю заполненную строку
            $highestRow = $sheet->getHighestRow();
            // определяем самую крайнюю заполненную колонку
            $highestColumn = $sheet->getHighestColumn();

            // флаг, который определяет начало данных
            $startData = false;
            // цикл разбора, от стартовой строки до нижней
            for ($row = 0; $row <= $highestRow; $row++) {
                // получаем массив данных строки
                $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
                $rowData = $rowData[0];

                // если очередная строка является заголовком категории цен
                if (trim($rowData[1]) == 'Square Hollow Sections') {
                    // ставим флаг в true
                    $startData = true;
                    // переходим к следующей итерации, где уже непосредственно будет работа с данными
                    continue;
                }

                if ($startData && ! trim($rowData[1])) { // цены по этой категории закончились
                    break;
                }

                // если данные еще не начались либо попалась пустая строка
                if (! $startData || ! trim($rowData[1]) || trim($rowData[1]) == 'Rectangular Hollow Sections') {
                    // переходим к следующей итерации
                    continue;
                }

                // размер
                $size = trim($rowData[1]);
                // цена за сталь S355J2(H)
                $steel = trim($rowData[2]);

                // кладем цену в массив
                $x["{$size}xEN10210xS355J2"] = $steel;
                $x["{$size}xEN10210xS355J2H"] = $steel;
            }
        }

        $data['HOLLOW SECTIONS'] = $x;

        // ********************************************************************
        // end HOLLOW SECTIONS
        // ********************************************************************

        // ********************************************************************
        // start FLAT BARS
        // ********************************************************************

        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('Flat Bars');
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
            if (trim($rowData[2]) == 'S235JR') {
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

            // убираем лишние пробелы и преобразуем к нижнему регистру
            $s = strtolower(str_replace(' ', '', $rowData[1]));
            // разбиваем строку на две части
            list($a, $b) = explode('x', $s);

            // в левой части ширина, диапазон либо перечень
            if (strpos($a, '-') !== false) {
                list($from, $to) = explode('-', $a);
                $widthes = range($from, $to);
            } elseif (strpos($a, ',') !== false) {
                $widthes = explode(',', $a);
            } else {
                $widthes = [$a];
            }

            // в правой части высота, диапазон либо перечень
            if (strpos($b, '-') !== false) {
                list($from, $to) = explode('-', $b);
                $heights = range($from, $to);
            } elseif (strpos($b, ',') !== false) {
                $heights = explode(',', $b);
            } else {
                $heights = [$b];
            }

            // цена за сталь S235JR, S235JR/S275JR, S275JR
            $steel1 = trim($rowData[2]);
            // цена за сталь S355J2, ST52-3
            $steel2 = trim($rowData[3]);

            // создаем специальный хеш для хранения цен
            foreach ($widthes as $w) {
                foreach ($heights as $h) {
                    $x["{$w}x{$h}xS235JR"] = $steel1;
                    $x["{$w}x{$h}xS235JR/S275JR"] = $steel1;
                    $x["{$w}x{$h}xS275JR"] = $steel1;
                    $x["{$w}x{$h}xS355J2"] = $steel2;
                    $x["{$w}x{$h}xST52-3"] = $steel2;
                }
            }
        }

        $data['FLAT BARS'] = $x;

        // ********************************************************************
        // end FLAT BARS
        // ********************************************************************

        // ********************************************************************
        // start ROUND BARS
        // ********************************************************************

        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('Round Bars');
        // определяем самую нижнюю заполненную строку
        $highestRow = $sheet->getHighestRow();
        // определяем самую крайнюю заполненную колонку
        $highestColumn = $sheet->getHighestColumn();

        $x = [];

        // флаг, который определяет начало данных
        $startData = false;
        // первый проход листа
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
//            if (trim($rowData[0]) == 'DRAWN BARS' && trim($rowData[1]) == 'S235JR') {
            if (trim($rowData[1]) == 'S235JR') {
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

            // цена за сталь S235JR, S235JR/S275JR
            $steel1 = trim($rowData[1]);
            // цена за сталь S355J2
            $steel2 = trim($rowData[2]);
            // цена за сталь C45, C45E
            $steel3 = trim($rowData[3]);
            // цена за сталь 40H
            $steel4 = trim($rowData[4]);

            // создаем специальный хеш для хранения цен
            foreach ($diameters as $d) {
                $x["{$d}xS235JR"] = $steel1;
                $x["{$d}xS235JR/S275JR"] = $steel1;
                $x["{$d}xS355J2"] = $steel2;
                $x["{$d}xC45"] = $steel3;
                $x["{$d}xC45E"] = $steel3;
                $x["{$d}x40H"] = $steel4;
            }
        }

        $data['ROUND BARS'] = $x;

        $x = [];
        $startData = false;
        // второй проход листа
        for ($row = 0; $row <= $highestRow; $row++) {
            // получаем массив данных строки
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
            $rowData = $rowData[0];

            // если очередная строка является заголовком категории цен
//            if (trim($rowData[1]) == '11SMn30') {
            if (trim($rowData[0]) == 'DRAWN ROUND BARS' && trim($rowData[1]) == 'S235JR') {
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

            //--------------
            // цена за сталь S235JR
            $steel1 = trim($rowData[1]);
            // цена за сталь S355J2
            $steel2 = trim($rowData[2]);
            // цена за сталь C45
            $steel3 = trim($rowData[3]);
            // цена за сталь 11SMn30
            $steel4 = trim($rowData[4]);

            // создаем специальный хеш для хранения цен
            foreach ($diameters as $d) {
                $x["{$d}xS235JR"] = $steel1;
                $x["{$d}xS355J2"] = $steel2;
                $x["{$d}xC45"] = $steel3;
                $x["{$d}x11SMn30"] = $steel4;
            }
            //--------------

            /*            // цена за сталь 11SMn30
                        $steel1 = trim($rowData[1]);
                        // цена за сталь 34HGS
                        $steel2 = trim($rowData[2]);

            // создаем специальный хеш для хранения цен
            foreach ($diameters as $d) {
                $x["{$d}x11SMn30"] = $steel1;
                $x["{$d}x34HGS"] = $steel2;
            }*/
        }

        $data['DRAWN ROUND BARS'] = $x;

        // ********************************************************************
        // end ROUND BARS
        // ********************************************************************

        // ********************************************************************
        // start SQUARE BARS
        // ********************************************************************

        // устанавливаем рабочий лист
        $sheet = $objPHPExcel->getSheetByName('Square Bars');
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
            if (trim($rowData[2]) == 'S235JR') {
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

            // убираем лишние пробелы
            $s = str_replace(' ', '', $rowData[1]);

            // получаем массив ширин
            if (strpos($s, '-') !== false) {
                list($from, $to) = explode('-', $s);
                $heights = range($from, $to);
            } else {
                $heights = [$s];
            }

            // цена за сталь S235JR, S235JR/S275JR, S275JR
            $steel1 = trim($rowData[2]);
            // цена за сталь S355J2
            $steel2 = trim($rowData[3]);

            // создаем специальный хеш для хранения цен
            foreach ($heights as $h) {
                $x["{$h}xS235JR"] = $steel1;
                $x["{$h}xS275JR"] = $steel1;
                $x["{$h}xS235JR/S275JR"] = $steel1;
                $x["{$h}xS355J2"] = $steel2;
            }
        }

        $data['SQUARE BARS'] = $x;

        // ********************************************************************
        // end SQUARE BARS
        // ********************************************************************

        // ***************
        // start RIBBED BARS & WIRE ROD
        // ***************
        $data['RIBBED BARS'] = \App\PriceParser::getRibberBars($objPHPExcel);
        $data['WIRE ROD'] = \App\PriceParser::getWireRod($objPHPExcel);

        // ***************
        // end RIBBED BARS & WIRE ROD
        // ***************
    }

    return $data;
}

/**
 * функция для разбора отстатков склада
 *
 * принимает на вход путь к файлу склада и настройки разбора
 * производит разбор данных и отдает массив, поделенный по складам
 *
 * string        stockFileName        путь к файлу склада
 * object    cfg                    объект с настройками
 *
 * return    array                разобранные данные
 */
function parseData($stockFileName, $cfg) {
    // подключение утилиты для разбора XLS
    //require_once(dirname(__FILE__) . '/vendor/PHPExcel/IOFactory.php');

    // флаг для определения, нужно ли продолжать разбор
    $processing = true;

    try { // попытка открыть файл на чтение
        $inputFileType = PHPExcel_IOFactory::identify($stockFileName);
        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objPHPExcel = $objReader->load($stockFileName);
    } catch (Exception $e) { // что-то пошло не так (напирмер, битый файл)
        echo "error loading file - " . $e->getMessage();
        $processing = false; // указываем, что продолжать не нужно
    }

    if ($processing) { // файл открылся на чтение, можно разбирать
        // здесь будут храниться разобранные данные товаров
        $data = [];

        // получаем количество листов (каждый лист - склад)
        $sheetCount = $objPHPExcel->getSheetCount();

        for ($i = 0; $i < $sheetCount; $i++) { // обходим каждый лист
            // устанавливаем рабочий лист
            $sheet = $objPHPExcel->getSheet($i);
            // определяем самую нижнюю заполненную строку
            $highestRow = $sheet->getHighestRow();
            // определяем самую крайнюю заполненную колонку
            $highestColumn = $sheet->getHighestColumn();
            // имя текущего склада
            $stockName = $sheet->getTitle();
            // массив данных текущего склада
            $data[$stockName] = [];

            $header = [];
            // флаг, который определяет начало данных
            $startData = false;
            // цикл разбора, от стартовой строки до нижней
            for ($row = 0; $row <= $highestRow; $row++) {
                // получаем массив данных строки
                $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
                $rowData = $rowData[0];

                // если очередная строка является заголовком таблицы данных
                if (in_array(trim($rowData[0]), ['SPECIFIC NAME', 'Nazwa materiału'])) {
                    if (trim($rowData[0]) == 'Nazwa materiału') {
//                        $cfg->titles['unit'] = 2;
//                        $cfg->titles['quantity'] = 3;
                        $cfg->titles['quantity'] = 2;
                        $cfg->titles['unit'] = 3;
                    }
                    // ставим флаг в true
                    $startData = true;
                    // переходим к следующей итерации, где уже непосредственно будет работа с данными
                    continue;
                }

                // если данные еще не начались либо попалась пустая строка
                if (! $startData || ! trim($rowData[0])) {
                    // переходим к следующей итерации
                    continue;
                }

                // разбираем строку и добавляем полученные данные в массив склада
                $data[$stockName][] = parseRow($rowData, $cfg);
            }
        }

        return $data;
    }
}

/**
 * функция для вывода данных в файл XLS
 *
 * принимает на вход имя файла для записи в него данных и
 * массив разобранных данных. сохраняет все в файл XLS
 *
 * string        fileName        имя файла
 * array        data            массив данных
 *
 * return    void
 */
function saveDataToExcel($fileName, $data) {
    // подключение утилиты для записи в XLS
    //require_once(dirname(__FILE__) . '/vendor/PHPExcel/IOFactory.php');
    // создание экземпляра класса для работы с электронной таблицей
    $objPHPExcel = new PHPExcel();

    $i = 0;
    // каждый склад на своем листе
    foreach ($data as $stock => $products) {
        // создаем рабочий лист при необходимости
        if ($i) {
            $objPHPExcel->createSheet();
        }
        // устанавливаем активный лист
        $objPHPExcel->setActiveSheetIndex($i);

        // создаем шапку таблицы
        $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Оригинал');
        $objPHPExcel->getActiveSheet()->SetCellValue('B1', 'PL');
        $objPHPExcel->getActiveSheet()->SetCellValue('C1', 'EN');
        $objPHPExcel->getActiveSheet()->SetCellValue('D1', 'RU');
        $objPHPExcel->getActiveSheet()->SetCellValue('E1', 'Вид');
        $objPHPExcel->getActiveSheet()->SetCellValue('F1', 'Размеры');
        $objPHPExcel->getActiveSheet()->SetCellValue('G1', 'Толщина');
        $objPHPExcel->getActiveSheet()->SetCellValue('H1', 'Ширина');
        $objPHPExcel->getActiveSheet()->SetCellValue('I1', 'Высота');
        $objPHPExcel->getActiveSheet()->SetCellValue('J1', 'Длина');
        $objPHPExcel->getActiveSheet()->SetCellValue('K1', 'Диаметр (мм)');
        $objPHPExcel->getActiveSheet()->SetCellValue('L1', 'Диаметр (дюймы)');
        $objPHPExcel->getActiveSheet()->SetCellValue('M1', 'Толщина стенки');
        $objPHPExcel->getActiveSheet()->SetCellValue('N1', 'Сталь');
        $objPHPExcel->getActiveSheet()->SetCellValue('O1', 'Обработка');
        $objPHPExcel->getActiveSheet()->SetCellValue('P1', 'Примечание');
        $objPHPExcel->getActiveSheet()->SetCellValue('Q1', 'Количество');
        $objPHPExcel->getActiveSheet()->SetCellValue('R1', 'Цена');
        $objPHPExcel->getActiveSheet()->SetCellValue('S1', 'Наименование в 1С');
        $objPHPExcel->getActiveSheet()->SetCellValue('T1', 'Вид номенклатуры');

        // обходим все товары и переносим их на лист
        $j = 2;
        foreach ($products as $p) {
            $objPHPExcel->getActiveSheet()->SetCellValue('A' . $j, (isset($p->orig) ? $p->orig : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('B' . $j, (isset($p->namePL) ? $p->namePL : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('C' . $j, (isset($p->nameEN) ? $p->nameEN : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('D' . $j, (isset($p->nameRU) ? $p->nameRU : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('E' . $j, (isset($p->type) ? $p->type : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('F' . $j, (isset($p->size) ? $p->size : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('G' . $j, (isset($p->thickness) ? $p->thickness : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('H' . $j, (isset($p->width) ? $p->width : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('I' . $j, (isset($p->height) ? $p->height : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('J' . $j, (isset($p->length) ? $p->length : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('K' . $j, (isset($p->diameter) ? $p->diameter : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('L' . $j, (isset($p->inches) ? $p->inches : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('M' . $j, (isset($p->wall) ? $p->wall : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('N' . $j, (isset($p->steel) ? $p->steel : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('O' . $j, (isset($p->processing) ? $p->processing : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('P' . $j, (isset($p->comment) ? $p->comment : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('Q' . $j, (isset($p->quantity) ? $p->quantity : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('R' . $j, (isset($p->price) ? $p->price : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('S' . $j, (isset($p->name1C) ? $p->name1C : ''));
            $objPHPExcel->getActiveSheet()->SetCellValue('T' . $j, (isset($p->nom) ? $p->nom : ''));

            $j++;
        }

        // именуем лист
        $objPHPExcel->getActiveSheet()->setTitle($stock);

        $i++;
    }

    // сохраняем файл
    $objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
    $objWriter->save($fileName);
}

/**
 * функция для вывода данных в файл XLS
 *
 * принимает на вход массив разобранных данных и конфиг
 * сохраняет по складам в YML
 *
 * array        data            массив данных
 * object    cfg            объект с настройками
 *
 * return    void
 */
function saveDataToYML($data) {
    $date = date('Y-m-d H:i');

    $i = 0;
    // каждый склад на своем листе
    foreach ($data as $stock => $products) {
        $xml = "<?xml version='1.0' encoding='UTF-8'?>" . PHP_EOL;
        $xml .= "<yml_catalog date='{$date}'>" . PHP_EOL;
        $xml .= "  <shop>" . PHP_EOL;
        $xml .= "    <company>STALPROFIL SA</company>" . PHP_EOL;
        $xml .= "    <company_id></company_id>" . PHP_EOL;
        $xml .= "    <sklad></sklad>" . PHP_EOL;
        $xml .= "    <sklad_name>{$stock}</sklad_name>" . PHP_EOL;
        $xml .= "    <offers>" . PHP_EOL;

        foreach ($products as $p) {
            $xml .= "      <offer>" . PHP_EOL;
            $xml .= "        <section_id></section_id>" . PHP_EOL; // todo fill map, get data
            $xml .= "        <price>" . (isset($p->price) ? $p->price : '') . "</price>" . PHP_EOL;
            $xml .= "        <measure>т</measure>" . PHP_EOL;
            $xml .= "        <quantity>" . (isset($p->quantity) ? $p->quantity : '') . "</quantity>" . PHP_EOL;
            $xml .= "        <name>" . (isset($p->nameRU) ? $p->nameRU : '') . "</name>" . PHP_EOL;
            $xml .= "        <properties>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Вид</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->type) ? $p->type : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Размеры</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->size) ? $p->size : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Толщина</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->thickness) ? $p->thickness : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Ширина</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->width) ? $p->width : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Высота</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->height) ? $p->height : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Диаметр (мм)</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->diameter) ? $p->diameter : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Диаметр (дюймы)</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->inches) ? $p->inches : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Толщина стенки</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->wall) ? $p->wall : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Сталь</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->steel) ? $p->steel : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Обработка</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->processing) ? $p->processing : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "          <property>" . PHP_EOL;
            $xml .= "            <id></id>" . PHP_EOL;
            $xml .= "            <name>Примечание</name>" . PHP_EOL;
            $xml .= "            <value>" . (isset($p->comment) ? $p->comment : '') . "</value>" . PHP_EOL;
            $xml .= "          </property>" . PHP_EOL;

            $xml .= "        </properties>" . PHP_EOL;
            $xml .= "      </offer>" . PHP_EOL;
        }

        $xml .= "    </offers>" . PHP_EOL;
        $xml .= "  </shop>" . PHP_EOL;
        $xml .= "</yml_catalog>";

        $i++;
        file_put_contents("stock_{$i}.yml", $xml);
    }
}

/**
 * функция для вывода данных в сыром виде
 *
 * принимает на вход массив разобранных данных и выводит его без обработки
 *
 * array    data        массив данных
 *
 * return    void
 */
function printRaw($data) {
    echo "<pre>";
    print_r($data);
    echo "</pre>";
}

/**
 * функция для вывода данных в виде таблицы
 *
 * принимает на вход массив разобранных данных и выводит его в виде таблицы
 *
 * array    data        массив данных
 *
 * return    void
 */
function printTable($data) {
    echo "<!DOCTYPE html><head></head><html><body>";
    echo "<style>table tr:not(.cap):hover {background:lime;}</style>";

    $i = 0;
    foreach ($data as $stock => $products) {
        echo "<a href='#stock{$i}'>{$stock}</a><br>";
        $i++;
    }

    $total = 0;
    $priceless = 0;

    $cap = "
		<tr class='cap'>
			<th>Оригинал</th>
			<th>PL</th>
			<th>EN</th>
			<th>RU</th>
			<th>Вид</th>
			<th>Размеры</th>
			<th>Толщина</th>
			<th>Ширина</th>
			<th>Высота</th>
			<th>Длина</th>
			<th>Диаметр (мм)</th>
			<th>Диаметр (дюймы)</th>
			<th>Толщина стенки</th>
			<th>Сталь</th>
			<th>Обработка</th>
			<th>Примечание</th>
			<th>Количество</th>
			<th>Цена</th>
			<th>Наименование в 1С</th>
			<th>Вид номенклатуры</th>
		</tr>
	";

    $i = 0;
    foreach ($data as $stock => $products) {
        echo "<a name='stock{$i}'>{$stock}</a>";
        echo "<table border='1' style='border-collapse:collapse'>";
        echo $cap;
        $j = 0;
        foreach ($products as $p) {
            $total++;
            if ($p->price == 0) {
                $priceless++;
            }
            echo "
				<tr>
					<td>" . (isset($p->orig) ? $p->orig : '') . "</td>
					<td>" . (isset($p->namePL) ? $p->namePL : '') . "</td>
					<td>" . (isset($p->nameEN) ? $p->nameEN : '') . "</td>
					<td>" . (isset($p->nameRU) ? $p->nameRU : '') . "</td>
					<td>" . (isset($p->type) ? $p->type : '') . "</td>
					<td>" . (isset($p->size) ? $p->size : '') . "</td>
					<td>" . (isset($p->thickness) ? $p->thickness : '') . "</td>
					<td>" . (isset($p->width) ? $p->width : '') . "</td>
					<td>" . (isset($p->height) ? $p->height : '') . "</td>
					<td>" . (isset($p->length) ? $p->length : '') . "</td>
					<td>" . (isset($p->diameter) ? $p->diameter : '') . "</td>
					<td>" . (isset($p->inches) ? $p->inches : '') . "</td>
					<td>" . (isset($p->wall) ? $p->wall : '') . "</td>
					<td>" . (isset($p->steel) ? $p->steel : '') . "</td>
					<td>" . (isset($p->processing) ? $p->processing : '') . "</td>
					<td>" . (isset($p->comment) ? $p->comment : '') . "</td>
					<td>" . (isset($p->quantity) ? $p->quantity : '') . "</td>
					<td " . (($p->price == 0) ? "style='background:orange'" : '') . ">" . (isset($p->price) ? $p->price : '') . "</td>
					<td>" . (isset($p->name1C) ? $p->name1C : '') . "</td>
					<td>" . (isset($p->nom) ? $p->nom : '') . "</td>
				</tr>
			";
            if (++$j % 50 == 0) {
                echo $cap;
            }
        }
        echo "</table>";
        $i++;
    }

    $percentage = number_format(100 - ($priceless / $total * 100), 2);
    echo "Всего товаров = {$total}<br>";
    echo "Товаров без цены = {$priceless}<br>";
    echo "Процент обработанных товаров = {$percentage}%<br>";
    echo "</body></html>";
}

/**
 * функция для вывода результата работы парсера
 *
 * принимает на вход массив разобранных данных и возвращает результаты разбора
 *
 * array        data        массив данных
 *
 * return    string
 */
function parseResult($data) {
    $ret = '';

    $total = 0;
    $priceless = 0;

    foreach ($data as $stock => $products) {
        foreach ($products as $p) {
            $total++;
            if ($p->price == 0) {
                $priceless++;
            }
        }
    }

    $percentage = number_format(100 - ($priceless / $total * 100), 2);
    $ret .= "Всего товаров = {$total}<br>";
    $ret .= "Товаров без цены = {$priceless}<br>";
    $ret .= "Процент обработанных товаров = {$percentage}%";

    return $ret;
}

/**
 * функция для обработки файла
 *
 * получает на вход ключ для суперглобального массива,
 * сохраняет файл при его наличии или отдает false при наличии ошибок.
 *
 * string        key        массив данных
 *
 * return    false
 */
function processFile($key) {
    if (! isset($_FILES[$key]) || empty($_FILES[$key]) || $_FILES[$key]['error']) {
        return false;
    }

    move_uploaded_file($_FILES[$key]['tmp_name'], "{$key}.xls");

    return true;
}

/**
 * функция для обработки формы
 *
 * разбирает поля формы, сохраняет файлы и записывает результат в сессию
 *
 * return    void
 */
function processForm($parserConfiguration) {
    global $prices;

    $priceFile = processFile('price');
    $stockFile = processFile('stock');

    $canProcess = true;

    if (false === $priceFile) {
        $_SESSION['errors'][] = 'Вы забыли прикрепить файл с ценами';
        $canProcess = false;
    }

    if (false === $stockFile) {
        $_SESSION['errors'][] = 'Вы забыли прикрепить файл с товарами';
        $canProcess = false;
    }

    if ($canProcess) {
        $prices = parsePrices('price.xls', $parserConfiguration); // разбор цен
        $data = parseData('stock.xls', $parserConfiguration); // разбор складов
        saveDataToExcel('output.xls', $data); // сохранение данных в файл XLS
        $_SESSION['success'][] = parseResult($data);
    }

    header('Location: index.php');
}
