"""
JSA Skill 验证工具
检查 skill 结构完整性和文档规范

使用方法:
    python validate_skill.py
"""

import os
import json

def validate_skill():
    """验证 skill 结构完整性"""
    
    skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    print("=" * 50)
    print("JSA-vibcoding Skill 验证")
    print("=" * 50)
    
    errors = []
    warnings = []
    
    # 1. 检查必需文件
    required_files = [
        "SKILL.md",
        "docs/API参考.md",
        "examples/basic_usage.js",
        "references/best_practices.md",
        "references/code_templates.md",
        "references/enum_constants.md",
        "references/jsa_patterns.md"
    ]
    
    print("\n[检查必需文件]")
    for file in required_files:
        path = os.path.join(skill_dir, file)
        if os.path.exists(path):
            size = os.path.getsize(path)
            print(f"  ✓ {file} ({size} bytes)")
        else:
            print(f"  ✗ {file} (缺失)")
            errors.append(f"缺少文件: {file}")
    
    # 2. 检查 SKILL.md 格式
    print("\n[检查 SKILL.md 格式]")
    skill_md_path = os.path.join(skill_dir, "SKILL.md")
    if os.path.exists(skill_md_path):
        with open(skill_md_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 检查 frontmatter
        if content.startswith('---'):
            print("  ✓ 包含 frontmatter")
            
            # 提取 frontmatter
            end_idx = content.find('---', 3)
            if end_idx > 0:
                frontmatter = content[3:end_idx].strip()
                try:
                    # 简单解析
                    lines = frontmatter.split('\n')
                    for line in lines:
                        if ':' in line:
                            key, value = line.split(':', 1)
                            print(f"    - {key.strip()}: {value.strip()}")
                except:
                    warnings.append("frontmatter 格式可能有问题")
        else:
            warnings.append("缺少 frontmatter")
        
        # 检查必需章节
        required_sections = [
            "## 快速开始",
            "## JSA 代码规范",
            "## 常用API速查",
            "## 知识库导航"
        ]
        
        for section in required_sections:
            if section in content:
                print(f"  ✓ 包含章节: {section}")
            else:
                warnings.append(f"缺少章节: {section}")
    
    # 3. 检查文档大小
    print("\n[检查文档大小]")
    for root, dirs, files in os.walk(skill_dir):
        for file in files:
            if file.endswith('.md') or file.endswith('.js'):
                path = os.path.join(root, file)
                size = os.path.getsize(path)
                rel_path = os.path.relpath(path, skill_dir)
                
                if size > 50000:
                    warnings.append(f"文件过大 ({size} bytes): {rel_path}")
                elif size < 100:
                    warnings.append(f"文件过小 ({size} bytes): {rel_path}")
    
    # 4. 检查目录结构
    print("\n[检查目录结构]")
    required_dirs = ["docs", "examples", "references", "tools"]
    for dir_name in required_dirs:
        path = os.path.join(skill_dir, dir_name)
        if os.path.isdir(path):
            file_count = len([f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))])
            print(f"  ✓ {dir_name}/ ({file_count} 文件)")
        else:
            print(f"  ✗ {dir_name}/ (缺失)")
            errors.append(f"缺少目录: {dir_name}")
    
    # 5. 输出结果
    print("\n" + "=" * 50)
    print("验证结果")
    print("=" * 50)
    
    if errors:
        print(f"\n❌ 错误 ({len(errors)}):")
        for error in errors:
            print(f"  - {error}")
    
    if warnings:
        print(f"\n⚠️  警告 ({len(warnings)}):")
        for warning in warnings:
            print(f"  - {warning}")
    
    if not errors and not warnings:
        print("\n✅ Skill 结构完整，无问题！")
    elif not errors:
        print(f"\n✅ Skill 结构完整，有 {len(warnings)} 个警告")
    else:
        print(f"\n❌ Skill 有 {len(errors)} 个错误需要修复")
    
    return len(errors) == 0


if __name__ == "__main__":
    success = validate_skill()
    exit(0 if success else 1)
