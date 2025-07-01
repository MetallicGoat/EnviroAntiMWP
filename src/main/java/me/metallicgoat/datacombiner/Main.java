package me.metallicgoat.datacombiner;


import javafx.application.Application;
import javafx.stage.Stage;
import me.metallicgoat.datacombiner.util.Util;

public class Main extends Application {

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage stage) {
        final UserInterface userInterface = new UserInterface();

        userInterface.start(stage);

        Util.initUserInterface(userInterface);
    }
}
