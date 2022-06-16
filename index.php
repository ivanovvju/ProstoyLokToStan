<?php

/**
 * Парсинг выгруженной excel роботом и формирование справки "Простой локомотивов на станции более 5 часов (от прибытия до отправления)".
 * @author Иванов В.Ю.
 * @copyright 2022
 */

require_once 'Log.php';
require_once 'App.php';

echo "Start program";
Log::Info("Start program");

$app = new App();
if ($app->run()) {
    echo "End program";
    Log::Info("End program");
}
