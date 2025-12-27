"""
Data models for HAZOP analysis nodes and deviations.
"""
from dataclasses import dataclass, field, asdict
from typing import List, Optional
import json


@dataclass
class Deviation:
    """Represents a HAZOP deviation/note attached to a node."""
    deviation: str = ""
    causes: List[str] = field(default_factory=list)
    consequence: str = ""
    safeguards: List[str] = field(default_factory=list)
    recommendations: List[str] = field(default_factory=list)
    comments: str = ""
    minimized: bool = False
    
    def to_dict(self):
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data):
        return cls(**data)


@dataclass
class Node:
    """Represents a line/node in the HAZOP analysis."""
    name: str = ""
    color: str = "#FF0000"  # Default red
    thickness: int = 2
    transparency: float = 0.7
    has_arrow: bool = True
    font_size: int = 12
    points: List[tuple] = field(default_factory=list)  # List of (x, y) coordinates
    deviations: List[Deviation] = field(default_factory=list)
    page_number: int = 0
    
    def to_dict(self):
        return {
            'name': self.name,
            'color': self.color,
            'thickness': self.thickness,
            'transparency': self.transparency,
            'has_arrow': self.has_arrow,
            'font_size': self.font_size,
            'points': self.points,
            'deviations': [d.to_dict() for d in self.deviations],
            'page_number': self.page_number
        }
    
    @classmethod
    def from_dict(cls, data):
        node = cls(
            name=data.get('name', ''),
            color=data.get('color', '#FF0000'),
            thickness=data.get('thickness', 2),
            transparency=data.get('transparency', 0.7),
            has_arrow=data.get('has_arrow', True),
            font_size=data.get('font_size', 12),
            points=data.get('points', []),
            page_number=data.get('page_number', 0)
        )
        node.deviations = [Deviation.from_dict(d) for d in data.get('deviations', [])]
        return node


class HAZOPData:
    """Manages all HAZOP analysis data."""
    def __init__(self, pdf_path: str = ""):
        self.pdf_path = pdf_path
        self.nodes: List[Node] = []
    
    def add_node(self, node: Node):
        self.nodes.append(node)
    
    def remove_node(self, node: Node):
        if node in self.nodes:
            self.nodes.remove(node)
    
    def get_nodes_for_page(self, page_number: int) -> List[Node]:
        return [node for node in self.nodes if node.page_number == page_number]
    
    def to_dict(self):
        return {
            'pdf_path': self.pdf_path,
            'nodes': [node.to_dict() for node in self.nodes]
        }
    
    def to_json(self, filepath: str):
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
    
    @classmethod
    def from_dict(cls, data):
        hazop_data = cls(pdf_path=data.get('pdf_path', ''))
        hazop_data.nodes = [Node.from_dict(n) for n in data.get('nodes', [])]
        return hazop_data
    
    @classmethod
    def from_json(cls, filepath: str):
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return cls.from_dict(data)

