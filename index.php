<?php

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

session_start();
$errors = (isset($_SESSION) && isset($_SESSION['errors'])) ? $_SESSION['errors'] : array();
$success = (isset($_SESSION) && isset($_SESSION['success'])) ? $_SESSION['success'] : array();
unset($_SESSION['errors']);
unset($_SESSION['success']);
?>
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <link rel="stylesheet" href="assets/css/bootstrap.min.css">
    <title>Stalprofil parser</title>
</head>
<body>
<div class="container">
    <div class="row">
        <div class="col-sm-12 text-center">
            <h1>Парсер прайсов Stalprofil</h1>
        </div>
        <div class="col-sm-6 offset-3 bg-light text-dark rounded border border-primary">
            <?php if ($errors) : ?>
                <div class="alert alert-danger" style="margin-top:10px" id="errorNotice">
                    <?php echo implode('<br>', $errors); ?>
                </div>
            <?php endif; ?>

            <?php if ($success) : ?>
                <div class="alert alert-success" style="margin-top:10px" id="successNotice">
                    <?php echo implode('<br>', $success); ?>
                    <br><a class="btn btn-primary" href="output.xls" target="_blank" download="output.xls">Скачать файл</a>
                </div>
            <?php endif; ?>

            <div class="alert alert-warning" style="display:none;margin-top:10px" id="processingNotice">
                Обработка в процессе
            </div>

            <form method="post" action="parser.php" enctype="multipart/form-data" id="parserForm">
                <div class="form-group">
                    <label for="priceFile">Прайс (файл с ценами)</label>
                    <input type="file" class="form-control-file" id="priceFile" name="price">
                </div>
                <div class="form-group">
                    <label for="stockFile">Склад (файл с товарами)</label>
                    <input type="file" class="form-control-file" id="stockFile" name="stock">
                </div>
                <div class="form-group text-center">
                    <button type="submit" class="btn btn-primary mb-2">Запустить парсер</button>
                </div>
            </form>
        </div>
    </div>
</div>
<script>
    (function () {
        var form             = document.getElementById('parserForm');
        var errorNotice      = document.getElementById('errorNotice');
        var successNotice    = document.getElementById('successNotice');
        var processingNotice = document.getElementById('processingNotice');
        form.onsubmit        = function () {
            if (null !== errorNotice) {
                errorNotice.style.display = 'none';
            }
            if (null !== successNotice) {
                successNotice.style.display = 'none';
            }
            processingNotice.style.display = 'block';
        };
    })();
</script>
</body>
</html>