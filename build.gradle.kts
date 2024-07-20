plugins {
    id("java")
}

group = "com.uniParse"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.poi:poi-ooxml:5.2.5")
    implementation("org.apache.poi:poi:5.2.5")
    implementation("org.projectlombok:lombok:1.18.30")
    implementation("org.jsoup:jsoup:1.17.2")
    testImplementation(platform("org.junit:junit-bom:5.9.1"))
    testImplementation("org.junit.jupiter:junit-jupiter")


}

tasks.test {
    useJUnitPlatform()
}