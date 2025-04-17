package io.projekt.utm_bot.service;

import io.projekt.utm_bot.config.BotConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Color;
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
    private final Map<Long, Byte> userWeekType = new HashMap<>();
    private static final String EXCEL_PATH = "src/main/resources/orar2.xlsx";

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

            if (text.startsWith("/schedule_for_the_day")) {
                userGroup.remove(chatId);
                userDay.remove(chatId);
                userWeekType.remove(chatId);
                sendMessage(chatId, "👋 Введи свою группу:");
            } else if (!userGroup.containsKey(chatId)) {
                userGroup.put(chatId, text.trim());
                sendDaySelection(chatId);
            } else {
                sendMessage(chatId, "❗ Пожалуйста, выбери день недели с помощью кнопок.");
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

                try {
                    String group = userGroup.get(chatId);
                    String day = userDay.get(chatId);
                    String schedule = extractScheduleByDay(EXCEL_PATH, group, day, week);
                    sendMessage(chatId, formatScheduleResponse(schedule, group, day));
                } catch (Exception e) {
                    sendMessage(chatId, "⚠ Ошибка при получении расписания: " + e.getMessage());
                }

                userDay.remove(chatId);
                userGroup.remove(chatId);
                userWeekType.remove(chatId);
            }
        }
    }

    private String formatScheduleResponse(String schedule, String group, String day) {
        return schedule == null || schedule.isBlank()
                ? "❌ Расписание не найдено для группы " + group + " на " + day + "."
                : "📅 Расписание для группы " + group + " на " + day + ":\n\n" + schedule;
    }

    private void sendDaySelection(long chatId) {
        String[] days = {"Luni", "Marţi", "Miercuri", "Joi", "Vineri", "Sâmbătă"};
        List<List<InlineKeyboardButton>> buttons = new ArrayList<>();
        for (String day : days) {
            InlineKeyboardButton btn = new InlineKeyboardButton(day);
            btn.setCallbackData("DAY_" + day);
            buttons.add(Collections.singletonList(btn));
        }
        InlineKeyboardMarkup markup = new InlineKeyboardMarkup();
        markup.setKeyboard(buttons);
        sendMessage(chatId, "📆 Выбери день недели:", markup);
    }

    private void sendWeekTypeSelection(long chatId) {
        InlineKeyboardButton odd = new InlineKeyboardButton("1 (Нечетная)");
        odd.setCallbackData("WEEK_1");

        InlineKeyboardButton even = new InlineKeyboardButton("2 (Четная)");
        even.setCallbackData("WEEK_2");

        List<List<InlineKeyboardButton>> rows = List.of(List.of(odd, even));
        InlineKeyboardMarkup markup = new InlineKeyboardMarkup();
        markup.setKeyboard(rows);
        sendMessage(chatId, "📖 Выбери тип недели:", markup);
    }

    private void sendMessage(long chatId, String text, InlineKeyboardMarkup markup) {
        SendMessage message = new SendMessage(String.valueOf(chatId), text);
        if (markup != null) message.setReplyMarkup(markup);
        try {
            execute(message);
        } catch (TelegramApiException e) {
            e.printStackTrace();
        }
    }

    private void sendMessage(long chatId, String text) {
        sendMessage(chatId, text, null);
    }

    public static String extractScheduleByDay(String path, String group, String selectedDay, byte isEvenWeek) throws IOException {
        FileInputStream fis = new FileInputStream(new File(path));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter formatter = new DataFormatter();
        StringBuilder schedule = new StringBuilder();

        int groupRowIndex = 8;
        int groupCol = -1;

        Row groupRow = sheet.getRow(groupRowIndex);
        for (int col = 0; col < groupRow.getLastCellNum(); col++) {
            String cellVal = formatter.formatCellValue(groupRow.getCell(col)).trim();
            if (cellVal.equalsIgnoreCase(group)) {
                groupCol = col;
                break;
            }
        }

        if (groupCol == -1) {
            workbook.close();
            return "❌ Группа \"" + group + "\" не найдена.";
        }

        int dayCol = 1;
        int dayRowStart = -1;
        for (int row = 0; row <= sheet.getLastRowNum(); row++) {
            Row r = sheet.getRow(row);
            if (r == null) continue;
            String cellVal = formatter.formatCellValue(r.getCell(dayCol)).trim();
            if (cellVal.equalsIgnoreCase(selectedDay)) {
                dayRowStart = row;
                break;
            }
        }

        if (dayRowStart == -1) {
            workbook.close();
            return "❌ День \"" + selectedDay + "\" не найден.";
        }

        List<String> allowedTimes = Arrays.asList(
                "8.00-9.30", "9.45-11.15", "11.30-13.00",
                "13.15-14.45", "15.00-16.30", "16.45-18.15", "18.30-20.00", "20.15"
        );

        for (int row = dayRowStart + 1; row <= sheet.getLastRowNum(); row++) {
            Row r = sheet.getRow(row);
            if (r == null) continue;

            String time = formatter.formatCellValue(r.getCell(2)).trim();
            if (time.isBlank() || !allowedTimes.contains(time)) continue;

            Cell lessonCell = r.getCell(groupCol);
            String lesson = formatter.formatCellValue(lessonCell).trim();

            boolean hasFill = false;
            if (lessonCell != null && lessonCell.getCellStyle() != null) {
                Color color = lessonCell.getCellStyle().getFillForegroundColorColor();
                hasFill = color != null;
            }

            if (!lesson.isBlank() || hasFill) {
                schedule.append("📚 ").append(lesson.isBlank() ? "(пусто, но залито)" : lesson).append("\n")
                        .append("⏰ ").append(time).append("\n\n");
            }
        }

        workbook.close();

        return schedule.isEmpty()
                ? "ℹ Занятий не найдено на выбранный день."
                : schedule.toString().trim();
    }
}
