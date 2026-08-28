package org.aspose.slides.foss;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.fail;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import org.junit.jupiter.api.Test;

/**
 * Proves that the test sources outside {@code src/test/java} were compiled and are on the
 * test classpath.
 *
 * <p>This project keeps its integration and conformance tests in {@code tests/integration}
 * and {@code tests/conformance}, which Maven does not know about: they are added as test
 * source roots by {@code build-helper-maven-plugin} in {@code pom.xml}. A root that is
 * missing from that configuration is not an error. It compiles nothing, runs nothing, and
 * reports BUILD SUCCESS - the whole suite disappears and the build stays green, which is
 * exactly the failure this project exists to catch in the files it writes.</p>
 *
 * <p>Removing {@code tests/conformance} from the plugin was tried, and the build failed only
 * by accident: one integration test happens to import a class from the conformance package.
 * Delete that one import and the conformance suite vanishes silently. So the guard is here
 * rather than left to chance, and it is here in {@code src/test/java} on purpose - that is
 * the one root Maven compiles without being told, so this test cannot disappear along with
 * the roots it is checking.</p>
 *
 * <p>It checks every {@code .java} file under {@code tests/}, not a hard-coded list, so a
 * new source root added there is covered the day it is added.</p>
 */
class TestSourceRootsTest {

    /** Matches the package declaration of a compilation unit. */
    private static final Pattern PACKAGE = Pattern.compile(
            "^\\s*package\\s+([A-Za-z0-9_.]+)\\s*;", Pattern.MULTILINE);

    /**
     * The roots that exist today. Listed so that deleting one is a failure rather than a
     * shorter, still-green run: a check that walks a directory tree passes trivially once
     * the tree is gone.
     */
    private static final List<String> KNOWN_ROOTS = List.of("integration", "conformance");

    private static Path testsDirectory() {
        Path tests = Path.of("").toAbsolutePath().resolve("tests");
        assertThat(tests)
                .as("the extra test source roots live in %s; surefire runs with the project "
                        + "directory as its working directory, so this path is where they are "
                        + "expected to be", tests)
                .isDirectory();
        return tests;
    }

    @Test
    void everyExtraTestSourceRootMustBeOnTheTestClasspath() throws IOException {
        Path tests = testsDirectory();

        List<String> notCompiled = new ArrayList<>();
        int checked = 0;

        try (Stream<Path> files = Files.walk(tests)) {
            for (Path file : files.filter(p -> p.toString().endsWith(".java")).toList()) {
                String source = Files.readString(file, StandardCharsets.UTF_8);
                Matcher matcher = PACKAGE.matcher(source);
                if (!matcher.find()) {
                    fail("%s declares no package, so the class it defines cannot be named", file);
                }
                String name = file.getFileName().toString();
                String className = matcher.group(1) + "."
                        + name.substring(0, name.length() - ".java".length());
                checked++;
                try {
                    Class.forName(className, false, getClass().getClassLoader());
                } catch (ClassNotFoundException e) {
                    notCompiled.add(className + "  (" + tests.relativize(file) + ")");
                }
            }
        }

        assertThat(notCompiled)
                .as("classes under %s that are not on the test classpath. Their source root is "
                        + "missing from the build-helper-maven-plugin configuration in pom.xml, "
                        + "so they were never compiled and never ran - and the build reported "
                        + "success anyway. %d of the %d sources found there did compile.",
                        tests, checked - notCompiled.size(), checked)
                .isEmpty();

        assertThat(checked)
                .as("sources found under %s; finding none would make this check pass without "
                        + "having checked anything", tests)
                .isPositive();
    }

    @Test
    void eachKnownTestSourceRootMustStillContributeTests() throws IOException {
        Path tests = testsDirectory();

        for (String root : KNOWN_ROOTS) {
            Path directory = tests.resolve(root);
            assertThat(directory)
                    .as("test source root %s. If it was removed on purpose, remove it from "
                            + "KNOWN_ROOTS here, from the build-helper configuration in pom.xml "
                            + "and from the suite table in CONTRIBUTING.md in the same change",
                            directory)
                    .isDirectory();

            List<String> loadedTests = new ArrayList<>();
            try (Stream<Path> files = Files.walk(directory)) {
                for (Path file : files.filter(p -> p.toString().endsWith("Test.java")).toList()) {
                    String source = Files.readString(file, StandardCharsets.UTF_8);
                    Matcher matcher = PACKAGE.matcher(source);
                    if (!matcher.find()) {
                        continue;
                    }
                    String name = file.getFileName().toString();
                    String className = matcher.group(1) + "."
                            + name.substring(0, name.length() - ".java".length());
                    try {
                        Class.forName(className, false, getClass().getClassLoader());
                        loadedTests.add(className);
                    } catch (ClassNotFoundException e) {
                        // Named by the other test, which reports the whole list at once.
                    }
                }
            }

            assertThat(loadedTests)
                    .as("test classes from %s that are on the test classpath", directory)
                    .isNotEmpty();
        }
    }
}
