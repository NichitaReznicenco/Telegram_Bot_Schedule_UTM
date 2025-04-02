package io.projekt.utm_bot.service;

import io.projekt.utm_bot.config.BotConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.telegram.telegrambots.bots.TelegramLongPollingBot;
import org.telegram.telegrambots.meta.api.methods.send.SendMessage;
import org.telegram.telegrambots.meta.api.objects.Update;
import org.telegram.telegrambots.meta.api.objects.replykeyboard.InlineKeyboardMarkup;
import org.telegram.telegrambots.meta.api.objects.replykeyboard.buttons.InlineKeyboardButton;
import org.telegram.telegrambots.meta.exceptions.TelegramApiException;

import java.io.*;
import java.util.*;

@Component
public class TelegramBot extends TelegramLongPollingBot {
    private final BotConfig config;
    private final Map<Long, String> userGroup = new HashMap<>();
    private final Map<Long, String> userDay = new HashMap<>();
    private final Map<Long, String> commandMap = new HashMap<>();
    private final Map<Long, Byte> userWeekType = new HashMap<>();

    public TelegramBot(BotConfig config) {
        this.config = config;
    }

    @Override
    public String getBotUsername() {
        return config.getBotName();
    }

    @Override
    public String getBotToken() {
        return config.getToken();
    }

    @Override
    public void clearWebhook() {}

    @Override
    public void onUpdateReceived(Update update) {
        if (update.hasMessage() && update.getMessage().hasText()) {
            String text = update.getMessage().getText();
            long chatId = update.getMessage().getChatId();

            if (commandMap.containsKey(chatId) && commandMap.get(chatId).equals("write_note")) {
                addNote(chatId, text);
                sendMessage(chatId, "✅ Заметка сохранена!");
                commandMap.remove(chatId);
                return;
            }

            switch (text.split(" ")[0]) {
                case "/schedule_for_the_day":
                    commandMap.put(chatId, "day");
                    userGroup.remove(chatId);
                    userDay.remove(chatId);
                    sendMessage(chatId, "Здравствуй, дорогой студент, ответь на пару вопросов и скажу тебе твое расписание. Первый вопрос, какая у тебя группа?");
                    break;

                case "/weekly_schedule":
                    commandMap.put(chatId, "week");
                    userGroup.remove(chatId);
                    sendMessage(chatId, "Введи свою группу, чтобы я показал тебе полное расписание на неделю.");
                    break;

                case "/my_notes":
                    List<String> notes = getNotes(chatId);
                    if (notes.isEmpty()) {
                        sendMessage(chatId, "🗒 У тебя пока нет заметок.");
                    } else {
                        StringBuilder response = new StringBuilder("📝 Твои заметки:\n");
                        for (int i = 0; i < notes.size(); i++) {
                            response.append(i + 1).append(") ").append(notes.get(i)).append("\n");
                        }
                        sendMessage(chatId, response.toString());
                    }
                    break;

                case "/note":
                    commandMap.put(chatId, "write_note");
                    sendMessage(chatId, "✍ Введи текст заметки:");
                    break;

                case "/delete_note":
                    try {
                        int index = Integer.parseInt(text.replace("/delete_note", "").trim()) - 1;
                        boolean success = deleteNote(chatId, index);
                        sendMessage(chatId, success ? "🗑 Заметка удалена!" : "❌ Неверный номер заметки.");
                    } catch (NumberFormatException e) {
                        sendMessage(chatId, "⚠ Неверный формат. Используй /delete_note <номер>");
                    }
                    break;

                default:
                    if (!userGroup.containsKey(chatId)) {
                        userGroup.put(chatId, text);
                        if ("day".equals(commandMap.get(chatId))) {
                            sendDaySelection(chatId);
                        } else if ("week".equals(commandMap.get(chatId))) {
                            sendWeekTypeSelection(chatId);
                        }
                    } else {
                        sendMessage(chatId, "Упсссс, ошибочка вышла");
                    }
                    break;
            }
        } else if (update.hasCallbackQuery()) {
            String data = update.getCallbackQuery().getData();
            long chatId = update.getCallbackQuery().getMessage().getChatId();

            if (data.startsWith("DAY_")) {
                userDay.put(chatId, data.replace("DAY_", ""));
                sendWeekTypeSelection(chatId);
            } else if (data.startsWith("WEEK_")) {
                byte week = Byte.parseByte(data.replace("WEEK_", ""));
                userWeekType.put(chatId, week);
                String group = userGroup.get(chatId);
                String cmd = commandMap.get(chatId);

                try {
                    if ("day".equals(cmd)) {
                        String day = userDay.get(chatId);
                        String schedule = extractScheduleByDay("src/main/resources/orar2.xlsx", group, day, week);
                        sendMessage(chatId, schedule == null || schedule.isBlank()
                                ? "❌ Расписание не найдено."
                                : "📅 Расписание для группы " + group + " на " + day + ":\n\n" + schedule);
                    } else if ("week".equals(cmd)) {
                        String schedule = extractWeeklySchedule("src/main/resources/orar2.xlsx", group, week);
                        sendMessage(chatId, schedule == null || schedule.isBlank()
                                ? "❌ Расписание не найдено."
                                : "🗓️ Полное недельное расписание для группы " + group + ":\n\n" + schedule);
                    }
                } catch (Exception e) {
                    sendMessage(chatId, "⚠ Произошла ошибка: " + e.getMessage());
                }

                userDay.remove(chatId);
                userGroup.remove(chatId);
                userWeekType.remove(chatId);
                commandMap.remove(chatId);
            }
        }
    }

    private void sendDaySelection(long chatId) {
        List<List<InlineKeyboardButton>> rows = new ArrayList<>();
        String[] days = {"Luni", "Marţi", "Miercuri", "Joi", "Vineri", "Sâmbătă"};
        for (String day : days) {
            InlineKeyboardButton button = new InlineKeyboardButton();
            button.setText(day);
            button.setCallbackData("DAY_" + day);
            rows.add(Collections.singletonList(button));
        }
        InlineKeyboardMarkup markup = new InlineKeyboardMarkup();
        markup.setKeyboard(rows);
        sendMessage(chatId, "Выберите день недели:", markup);
    }

    private void sendWeekTypeSelection(long chatId) {
        List<List<InlineKeyboardButton>> rows = new ArrayList<>();
        InlineKeyboardButton week1 = new InlineKeyboardButton();
        week1.setText("1 (Нечетная)");
        week1.setCallbackData("WEEK_1");

        InlineKeyboardButton week2 = new InlineKeyboardButton();
        week2.setText("2 (Четная)");
        week2.setCallbackData("WEEK_2");

        rows.add(Arrays.asList(week1, week2));
        InlineKeyboardMarkup markup = new InlineKeyboardMarkup();
        markup.setKeyboard(rows);
        sendMessage(chatId, "Выберите тип недели:", markup);
    }

    private void sendMessage(long chatId, String text, InlineKeyboardMarkup markup) {
        SendMessage message = new SendMessage();
        message.setChatId(String.valueOf(chatId));
        message.setText(text);
        if (markup != null) {
            message.setReplyMarkup(markup);
        }
        try {
            execute(message);
        } catch (TelegramApiException e) {
            e.printStackTrace();
        }
    }

    private void sendMessage(long chatId, String text) {
        sendMessage(chatId, text, null);
    }

    private void addNote(long chatId, String note) {
        File file = new File("src/main/resources/notes.txt");
        try (FileWriter writer = new FileWriter(file, true)) {
            writer.write(chatId + ":" + note + "\n");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private List<String> getNotes(long chatId) {
        List<String> notes = new ArrayList<>();
        File file = new File("src/main/resources/notes.txt");
        if (!file.exists()) return notes;

        try (Scanner scanner = new Scanner(file)) {
            while (scanner.hasNextLine()) {
                String line = scanner.nextLine();
                if (line.startsWith(chatId + ":")) {
                    notes.add(line.substring((chatId + ":").length()));
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return notes;
    }

    private boolean deleteNote(long chatId, int indexToDelete) {
        File inputFile = new File("src/main/resources/notes.txt");
        File tempFile = new File("src/main/resources/notes_temp.txt");
        boolean deleted = false;
        int currentIndex = 0;

        try (BufferedReader reader = new BufferedReader(new FileReader(inputFile));
             BufferedWriter writer = new BufferedWriter(new FileWriter(tempFile))) {

            String line;
            while ((line = reader.readLine()) != null) {
                if (line.startsWith(chatId + ":")) {
                    if (currentIndex == indexToDelete) {
                        deleted = true;
                        currentIndex++;
                        continue;
                    }
                    currentIndex++;
                }
                writer.write(line + "\n");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        if (!inputFile.delete() || !tempFile.renameTo(inputFile)) {
            return false;
        }

        return deleted;
    }

    public static String extractWeeklySchedule(String xlsxPath, String group, byte isEvenWeek) throws IOException {
        StringBuilder fullWeek = new StringBuilder();
        String[] days = {"Luni", "Marţi", "Miercuri", "Joi", "Vineri", "Sâmbătă"};
        for (String day : days) {
            String daily = extractScheduleByDay(xlsxPath, group, day, isEvenWeek);
            if (daily != null && !daily.isBlank()) {
                fullWeek.append("📅 ").append(day).append(":\n");
                fullWeek.append(daily).append("\n");
                fullWeek.append("━━━━━━━━━━━━━━━━━━━━━━\n");
            }
        }
        return fullWeek.toString();
    }

    public static String extractScheduleByDay(String xlsxPath, String group, String selectedDay, byte isEvenWeek) throws IOException {
        FileInputStream file = new FileInputStream(new File(xlsxPath));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter formatter = new DataFormatter();
        StringBuilder schedule = new StringBuilder();

        Row groupRow = sheet.getRow(0);
        int groupCol = -1;
        for (int col = 0; col < groupRow.getLastCellNum(); col++) {
            String cellVal = formatter.formatCellValue(groupRow.getCell(col));
            if (cellVal != null && cellVal.trim().equalsIgnoreCase(group)) {
                groupCol = col;
                break;
            }
        }
        if (groupCol == -1) {
            workbook.close();
            return null;
        }

        int startRow = -1;
        for (int row = 1; row <= sheet.getLastRowNum(); row++) {
            Row currentRow = sheet.getRow(row);
            if (currentRow == null) continue;
            String cellVal = formatter.formatCellValue(currentRow.getCell(0));
            if (cellVal != null && cellVal.trim().equalsIgnoreCase(selectedDay.trim())) {
                startRow = row;
                break;
            }
        }
        if (startRow == -1) {
            workbook.close();
            return null;
        }

        for (int row = startRow + 1; row <= startRow + 13 && row <= sheet.getLastRowNum(); row++) {
            Row currentRow = sheet.getRow(row);
            if (currentRow == null) continue;
            String time = formatter.formatCellValue(currentRow.getCell(1));
            if (time == null || time.isBlank()) continue;
            time = time.replace('.', ':').trim();

            Cell lessonCell = currentRow.getCell(groupCol);
            String lesson = (lessonCell != null) ? formatter.formatCellValue(lessonCell) : null;

            Row nextRow = (row + 1 <= sheet.getLastRowNum()) ? sheet.getRow(row + 1) : null;
            Cell altLessonCell = (nextRow != null) ? nextRow.getCell(groupCol) : null;
            String altLesson = (altLessonCell != null) ? formatter.formatCellValue(altLessonCell) : null;

            boolean isSplit = lesson != null && !lesson.isBlank() && altLesson != null && !altLesson.isBlank();

            if (isSplit) {
                if (isEvenWeek == 1) {
                    schedule.append("⏰ ").append(time).append("\n");
                    schedule.append("📚 ").append(lesson).append("\n\n");
                } else {
                    schedule.append("⏰ ").append(time).append("\n");
                    schedule.append("📚 ").append(altLesson).append("\n\n");
                }
                row++;
            } else if ((lesson != null && !lesson.isBlank()) || (altLesson != null && !altLesson.isBlank())) {
                String toPrint = (lesson != null && !lesson.isBlank()) ? lesson : altLesson;
                schedule.append("⏰ ").append(time).append("\n");
                schedule.append("📚 ").append(toPrint).append("\n\n");
            }
        }

        workbook.close();
        return schedule.toString().trim();
    }
}