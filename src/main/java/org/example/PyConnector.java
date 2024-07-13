package org.example;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

/**
 * Класс для запуска Python-скрипта из Java и обработки его вывода.
 */
public class PyConnector {
    private static final Logger logger = LoggerFactory.getLogger(PyConnector.class);

    /**
     * Запускает Python-скрипт и обрабатывает его вывод.
     * <p>
     * Этот метод запускает процесс Python, отправляет ему данные через стандартный ввод,
     * читает вывод и ошибки из процесса, а также ожидает его завершения.
     *
     * @param args аргументы командной строки (не используются)
     */
    public static void main(String[] args) {
        try {
            // Указываем путь к интерпретатору Python и скрипту
            Process process = getProcess();

            // Получаем стандартный ввод процесса (куда будем отправлять данные)
            OutputStream outputStream = process.getOutputStream();
            PrintWriter writer = new PrintWriter(outputStream, true); // true для автоочистки

            // Отправляем данные в Python-скрипт
            writer.println("Hello from Java");
            writer.flush();

            // Получаем стандартный вывод процесса (где будем читать результат)
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            BufferedReader errorReader = new BufferedReader(new InputStreamReader(process.getErrorStream()));
            String line;
            while ((line = reader.readLine()) != null) {
                logger.info("Вывод из Python: {}", line);
            }

            // Читаем ошибки из Python-скрипта (если есть)
            while ((line = errorReader.readLine()) != null) {
                logger.error("Ошибка из Python: {}", line);
            }

            // Ожидаем завершения процесса
            int exitCode = process.waitFor();
            logger.info("Процесс завершился с кодом: {}", exitCode);

            // Закрываем ресурсы
            reader.close();
            errorReader.close();
            writer.close();
        } catch (Exception e) {
            logger.error("Произошла ошибка", e);
        }
    }

    /**
     * Создает и запускает процесс для выполнения Python-скрипта.
     * <p>
     * Этот метод настраивает {@link ProcessBuilder} с указанными путями к интерпретатору Python и скрипту,
     * а затем запускает процесс.
     *
     * @return запущенный процесс
     * @throws IOException если возникает ошибка при создании процесса
     */
    private static Process getProcess() throws IOException {
        String pythonInterpreter = "C:\\Users\\Алексей\\AppData\\Local\\Programs\\Python\\Python312\\python.exe"; // Замените на путь к вашему интерпретатору Python
        String pythonScriptPath = "C:\\Users\\Алексей\\PycharmProjects\\EmergeData\\DateFormat.py"; // Замените на путь к вашему скрипту
        ProcessBuilder processBuilder = new ProcessBuilder(pythonInterpreter, pythonScriptPath);

        // Запускаем процесс
        return processBuilder.start();
    }

    /**
     * Вызов метода {@link #main(String[])} для запуска Python-скрипта.
     * <p>
     * Этот метод позволяет вызвать {@link #main(String[])} из других частей программы
     * с указанными аргументами командной строки.
     *
     * @param args аргументы командной строки (не используются)
     */
    public static void StartPyConnector(String[] args) {
        main(args);
    }
}
