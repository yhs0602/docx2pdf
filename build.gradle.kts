import org.jetbrains.compose.desktop.application.dsl.TargetFormat

plugins {
    kotlin("jvm")
    id("org.jetbrains.compose")
}

group = "com.yhs0602"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
    maven("https://maven.pkg.jetbrains.space/public/p/compose/dev")
    google()
}

dependencies {
    // Note, if you develop a library, you should use compose.desktop.common.
    // compose.desktop.currentOs should be used in launcher-sourceSet
    // (in a separate module for demo project and in testMain).
    // With compose.desktop.common you will also lose @Preview functionality
    implementation(compose.desktop.currentOs)
    // https://mvnrepository.com/artifact/org.apache.poi/poi
    implementation("org.apache.poi:poi:5.2.5")
    // https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml
    implementation("org.apache.poi:poi-ooxml:5.2.5")
    // https://mvnrepository.com/artifact/fr.opensagres.xdocreport/fr.opensagres.poi.xwpf.converter.pdf
    implementation("fr.opensagres.xdocreport:fr.opensagres.poi.xwpf.converter.pdf:2.0.6")
    // https://mvnrepository.com/artifact/fr.opensagres.xdocreport/org.odftoolkit.odfdom.converter.pdf
    implementation("fr.opensagres.xdocreport:org.odftoolkit.odfdom.converter.pdf:1.0.6")

}

compose.desktop {
    application {
        mainClass = "MainKt"

        nativeDistributions {
            targetFormats(TargetFormat.Dmg, TargetFormat.Msi, TargetFormat.Deb)
            packageName = "docx2pdf"
            packageVersion = "1.0.0"
        }
    }
}
