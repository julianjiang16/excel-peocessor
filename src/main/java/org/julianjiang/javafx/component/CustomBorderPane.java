package org.julianjiang.javafx.component;

import javafx.scene.image.Image;
import javafx.scene.layout.*;

public class CustomBorderPane extends BorderPane {

    public CustomBorderPane() {
        super();
        Image backgroundImage = new Image("bg/bg.png");
        BackgroundImage background = new BackgroundImage(backgroundImage, BackgroundRepeat.ROUND, BackgroundRepeat.ROUND, BackgroundPosition.DEFAULT, BackgroundSize.DEFAULT);
        this.setBackground(new Background(background));
    }
}
