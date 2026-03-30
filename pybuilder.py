import os
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


class PyBuilderApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PyBuilder - Python to EXE")
        self.root.geometry("860x560")
        self.root.minsize(780, 480)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.last_built_exe: Path | None = None
        self.is_building = False
        self.is_frozen = bool(getattr(sys, "frozen", False))

        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        top = ttk.LabelFrame(main, text="Файл для сборки", padding=10)
        top.pack(fill=tk.X)

        self.py_file_var = tk.StringVar()

        ttk.Label(top, text="Python файл (.py):").grid(row=0, column=0, sticky="w")

        self.path_entry = ttk.Entry(top, textvariable=self.py_file_var, width=90)
        self.path_entry.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(4, 0))

        browse_btn = ttk.Button(top, text="Выбрать...", command=self.pick_file)
        browse_btn.grid(row=1, column=1, pady=(4, 0))

        top.columnconfigure(0, weight=1)

        options = ttk.LabelFrame(main, text="Параметры", padding=10)
        options.pack(fill=tk.X, pady=(10, 0))

        self.one_file_var = tk.BooleanVar(value=True)
        self.auto_run_var = tk.BooleanVar(value=True)
        self.icon_file_var = tk.StringVar()
        self.version_var = tk.StringVar()

        ttk.Checkbutton(
            options,
            text="Собирать в один файл (--onefile)",
            variable=self.one_file_var,
        ).grid(row=0, column=0, sticky="w")

        ttk.Checkbutton(
            options,
            text="Автозапуск после сборки",
            variable=self.auto_run_var,
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        ttk.Label(
            options,
            text="Консоль: отключена (сборка всегда с --noconsole)",
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

        ttk.Label(options, text="Иконка EXE (.ico):").grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(options, textvariable=self.icon_file_var).grid(row=4, column=0, sticky="ew", padx=(0, 8), pady=(4, 0))
        ttk.Button(options, text="Иконка...", command=self.pick_icon).grid(row=4, column=1, pady=(4, 0))

        ttk.Label(options, text="Версия EXE (например 1.2.3.4):").grid(row=5, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(options, textvariable=self.version_var).grid(row=6, column=0, sticky="ew", padx=(0, 8), pady=(4, 0))

        options.columnconfigure(0, weight=1)

        buttons = ttk.Frame(main)
        buttons.pack(fill=tk.X, pady=(10, 0))

        self.build_btn = ttk.Button(buttons, text="Собрать EXE", command=self.start_build)
        self.build_btn.pack(side=tk.LEFT)

        self.run_btn = ttk.Button(
            buttons, text="Открыть собранное приложение", command=self.run_built_exe, state=tk.DISABLED
        )
        self.run_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.progress = ttk.Progressbar(buttons, mode="indeterminate")
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(12, 0))

        log_frame = ttk.LabelFrame(main, text="Лог", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=18)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)

        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def pick_file(self) -> None:
        picked = filedialog.askopenfilename(
            title="Выберите Python файл",
            filetypes=[("Python files", "*.py"), ("All files", "*.*")],
        )
        if picked:
            self.py_file_var.set(picked)

    def pick_icon(self) -> None:
        picked = filedialog.askopenfilename(
            title="Выберите ICO файл",
            filetypes=[("Icon files", "*.ico"), ("All files", "*.*")],
        )
        if picked:
            self.icon_file_var.set(picked)

    def start_build(self) -> None:
        if self.is_building:
            return

        src = self.py_file_var.get().strip().strip('"')
        if not src:
            messagebox.showwarning("Нет файла", "Сначала выберите Python файл.")
            return

        src_path = Path(src)
        if not src_path.exists() or src_path.suffix.lower() != ".py":
            messagebox.showerror("Ошибка", "Укажите существующий файл с расширением .py")
            return

        one_file = self.one_file_var.get()
        icon_raw = self.icon_file_var.get().strip().strip('"')
        icon_path: Path | None = None
        if icon_raw:
            icon_path = Path(icon_raw)
            if not icon_path.exists() or icon_path.suffix.lower() != ".ico":
                messagebox.showerror("Ошибка", "Иконка должна быть существующим .ico файлом.")
                return

        version_input = self.version_var.get().strip()
        normalized_version: str | None = None
        if version_input:
            normalized_version = self._normalize_version(version_input)
            if normalized_version is None:
                messagebox.showerror(
                    "Ошибка версии",
                    "Введите версию в формате чисел через точку, например: 1.0 или 1.2.3.4",
                )
                return

        self.last_built_exe = None
        self.run_btn.config(state=tk.DISABLED)
        self.is_building = True
        self.build_btn.config(state=tk.DISABLED)
        self.progress.start(10)

        self._log(f"Старт сборки: {src_path}")

        build_thread = threading.Thread(
            target=self._build_worker,
            args=(src_path, one_file, icon_path, normalized_version),
            daemon=True,
        )
        build_thread.start()

    def _build_worker(
        self,
        src_path: Path,
        one_file: bool,
        icon_path: Path | None,
        normalized_version: str | None,
    ) -> None:
        version_file_path: Path | None = None
        try:
            pyinstaller_cmd = self._find_pyinstaller_command()
            if pyinstaller_cmd is None and not self.is_frozen:
                self._log("PyInstaller не найден. Пробую установить через pip...")
                install_cmd = [sys.executable, "-m", "pip", "install", "pyinstaller"]
                install_result = subprocess.run(
                    install_cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    cwd=str(src_path.parent),
                    shell=False,
                )
                self._log(install_result.stdout)
                if install_result.returncode != 0:
                    self._log("Не удалось установить PyInstaller.")
                    self.root.after(0, self._build_finished, False, None)
                    return
                pyinstaller_cmd = self._find_pyinstaller_command()

            if pyinstaller_cmd is None:
                self._log("PyInstaller не найден в системе.")
                if self.is_frozen:
                    self._log("Для EXE-версии установите PyInstaller в PATH или через: py -m pip install pyinstaller")
                self.root.after(0, self._build_finished, False, None)
                return

            cmd = [*pyinstaller_cmd, "--clean", "--noconfirm"]
            if one_file:
                cmd.append("--onefile")
            cmd.append("--noconsole")
            if icon_path is not None:
                cmd.extend(["--icon", str(icon_path)])
            if normalized_version is not None:
                version_file_path = self._write_version_file(src_path, normalized_version)
                cmd.extend(["--version-file", str(version_file_path)])
            cmd.append(src_path.name)

            self._log(f"Команда: {' '.join(cmd)}")

            proc = subprocess.Popen(
                cmd,
                cwd=str(src_path.parent),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                shell=False,
            )

            assert proc.stdout is not None
            for line in proc.stdout:
                self._log(line.rstrip("\n"))

            return_code = proc.wait()
            if return_code != 0:
                self._log(f"Сборка завершилась с ошибкой. Код: {return_code}")
                self.root.after(0, self._build_finished, False, None)
                return

            built_exe = self._resolve_built_path(src_path, one_file)
            if built_exe is None:
                self._log("Сборка завершилась, но EXE не найден в папке dist.")
                self.root.after(0, self._build_finished, False, None)
                return

            self.root.after(0, self._build_finished, True, built_exe)
        except Exception as exc:
            self._log(f"Неожиданная ошибка: {exc}")
            self.root.after(0, self._build_finished, False, None)
        finally:
            if version_file_path and version_file_path.exists():
                try:
                    version_file_path.unlink()
                except OSError:
                    self._log(f"Не удалось удалить временный файл версии: {version_file_path}")

    def _find_pyinstaller_command(self) -> list[str] | None:
        candidates: list[list[str]] = []
        if self.is_frozen:
            candidates.extend([["py", "-m", "PyInstaller"], ["python", "-m", "PyInstaller"], ["pyinstaller"]])
        else:
            candidates.extend(
                [[sys.executable, "-m", "PyInstaller"], ["py", "-m", "PyInstaller"], ["pyinstaller"]]
            )

        seen: set[tuple[str, ...]] = set()
        for candidate in candidates:
            key = tuple(candidate)
            if key in seen:
                continue
            seen.add(key)

            check_target = candidate[0]
            if len(candidate) == 1 and shutil.which(check_target) is None:
                continue

            try:
                check = subprocess.run(
                    [*candidate, "--version"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    shell=False,
                )
            except FileNotFoundError:
                continue

            if check.returncode == 0:
                self._log(f"Найден PyInstaller: {' '.join(candidate)}")
                return candidate
        return None

    def _resolve_built_path(self, src_path: Path, one_file: bool) -> Path | None:
        dist_dir = src_path.parent / "dist"
        if not dist_dir.exists():
            return None

        name = src_path.stem

        if one_file:
            candidate = dist_dir / f"{name}.exe"
            if candidate.exists():
                return candidate
            return self._find_latest_dist_exe(dist_dir)

        folder_exe = dist_dir / name / f"{name}.exe"
        if folder_exe.exists():
            return folder_exe

        flat_exe = dist_dir / f"{name}.exe"
        if flat_exe.exists():
            return flat_exe

        return self._find_latest_dist_exe(dist_dir)

    def _find_latest_dist_exe(self, dist_dir: Path) -> Path | None:
        exe_files = sorted(dist_dir.rglob("*.exe"), key=lambda p: p.stat().st_mtime, reverse=True)
        return exe_files[0] if exe_files else None

    def _normalize_version(self, raw_value: str) -> str | None:
        parts = raw_value.strip().split(".")
        if not parts or len(parts) > 4:
            return None
        if any((not part.isdigit()) for part in parts):
            return None
        numbers = [str(int(part)) for part in parts]
        while len(numbers) < 4:
            numbers.append("0")
        return ".".join(numbers)

    def _write_version_file(self, src_path: Path, normalized_version: str) -> Path:
        comma_version = ", ".join(normalized_version.split("."))
        escaped_name = src_path.stem.replace("'", "")

        content = (
            "# UTF-8\n"
            "VSVersionInfo(\n"
            "  ffi=FixedFileInfo(\n"
            f"    filevers=({comma_version}),\n"
            f"    prodvers=({comma_version}),\n"
            "    mask=0x3F,\n"
            "    flags=0x0,\n"
            "    OS=0x40004,\n"
            "    fileType=0x1,\n"
            "    subtype=0x0,\n"
            "    date=(0, 0)\n"
            "    ),\n"
            "  kids=[\n"
            "    StringFileInfo(\n"
            "      [\n"
            "      StringTable(\n"
            "        '040904B0',\n"
            "        [StringStruct('CompanyName', 'PyBuilder'),\n"
            f"        StringStruct('FileDescription', '{escaped_name}'),\n"
            f"        StringStruct('FileVersion', '{normalized_version}'),\n"
            f"        StringStruct('InternalName', '{escaped_name}'),\n"
            f"        StringStruct('OriginalFilename', '{escaped_name}.exe'),\n"
            "        StringStruct('ProductName', 'PyBuilder Build'),\n"
            f"        StringStruct('ProductVersion', '{normalized_version}')])\n"
            "      ]),\n"
            "    VarFileInfo([VarStruct('Translation', [1033, 1200])])\n"
            "  ]\n"
            ")\n"
        )

        output_path = src_path.parent / "_pybuilder_version_info.txt"
        output_path.write_text(content, encoding="utf-8")
        self._log(f"Файл версии создан: {output_path}")
        return output_path

    def _build_finished(self, success: bool, exe_path: Path | None) -> None:
        self.is_building = False
        self.build_btn.config(state=tk.NORMAL)
        self.progress.stop()

        if success and exe_path is not None:
            self.last_built_exe = exe_path
            self.run_btn.config(state=tk.NORMAL)
            self._log(f"Успех! EXE: {exe_path}")
            if self.auto_run_var.get():
                try:
                    os.startfile(str(exe_path))
                    self._log(f"Автозапуск EXE: {exe_path}")
                    messagebox.showinfo(
                        "Готово",
                        f"Сборка завершена.\nФайл: {exe_path}\n\nПриложение запущено автоматически.",
                    )
                except Exception as exc:
                    self._log(f"Не удалось автозапустить EXE: {exc}")
                    messagebox.showwarning(
                        "Готово",
                        f"Сборка завершена.\nФайл: {exe_path}\n\nАвтозапуск не удался: {exc}",
                    )
            else:
                messagebox.showinfo("Готово", f"Сборка завершена.\nФайл: {exe_path}")
        else:
            messagebox.showerror("Ошибка", "Сборка не удалась. Проверьте лог.")

    def run_built_exe(self) -> None:
        if not self.last_built_exe or not self.last_built_exe.exists():
            messagebox.showwarning("Нет EXE", "Сначала соберите приложение.")
            self.run_btn.config(state=tk.DISABLED)
            return

        try:
            os.startfile(str(self.last_built_exe))
            self._log(f"Запуск EXE: {self.last_built_exe}")
        except Exception as exc:
            messagebox.showerror("Ошибка запуска", str(exc))

    def _log(self, text: str) -> None:
        self.log_queue.put(text)

    def _poll_log_queue(self) -> None:
        updated = False
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            else:
                updated = True
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, msg + "\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)

        if updated:
            self.root.update_idletasks()
        self.root.after(120, self._poll_log_queue)


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = PyBuilderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
